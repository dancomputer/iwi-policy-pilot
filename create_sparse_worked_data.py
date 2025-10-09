import pandas as pd
import numpy as np
import xarray as xr
import os
from pathlib import Path
import warnings
import re
warnings.filterwarnings('ignore')

def _find_planting_file(root_dir: str) -> str | None:
    for dp, _, fs in os.walk(root_dir):
        for fn in fs:
            low = fn.lower()
            if low.endswith((".nc", ".nc4")) and ("plant" in low or "pldate" in low):
                return os.path.join(dp, fn)
    return None

def _planting_start_year_from_folder(input_folder: str) -> int | None:
    """Return first calendar year present in planting dataset."""
    p = _find_planting_file(input_folder)
    if not p:
        print("[warn] planting file not found in", input_folder)
        return None
    ds = xr.open_dataset(p, decode_times=False)
    try:
        t = ds["time"]
        # If already datetime64[ns], use it directly
        if np.issubdtype(t.dtype, np.datetime64):
            y0 = pd.to_datetime(t.values[0]).year
            return int(y0)
        # else try to parse from units like 'days since 1982-01-01 ...'
        units = str(t.attrs.get("units", "")).lower()
        m = re.search(r"(\d{4})\s*-\s*\d{1,2}\s*-\s*\d{1,2}", units)
        return int(m.group(1)) + int(ds.time.values[0]) if m else None
    finally:
        ds.close()

def find_minimal_bounding_box(locations_file, grid_step=0.1):
    """
    Find minimal bounding box for worked locations with specified grid step
    """
    # Read the Excel file
    df = pd.read_excel(locations_file)
    
    # Get lat/lon bounds from worked locations
    min_lat = df['Latitude'].min()
    max_lat = df['Latitude'].max()
    min_lon = df['Longitude'].min()
    max_lon = df['Longitude'].max()
    
    # Define the global grid bounds
    global_lon_bounds = (-179.75, 179.75)
    global_lat_bounds = (-89.75, 89.75)
    
    # Calculate grid-aligned bounding box
    # Find the grid cells that contain our min/max points
    min_lon_grid = np.floor((min_lon - global_lon_bounds[0]) / grid_step) * grid_step + global_lon_bounds[0]
    max_lon_grid = np.ceil((max_lon - global_lon_bounds[0]) / grid_step) * grid_step + global_lon_bounds[0]
    min_lat_grid = np.floor((min_lat - global_lat_bounds[0]) / grid_step) * grid_step + global_lat_bounds[0]
    max_lat_grid = np.ceil((max_lat - global_lat_bounds[0]) / grid_step) * grid_step + global_lat_bounds[0]
    
    return {
        'min_lon': min_lon_grid,
        'max_lon': max_lon_grid,
        'min_lat': min_lat_grid,
        'max_lat': max_lat_grid
    }

def get_worked_locations_mask(locations_file, lat_coords, lon_coords, grid_step=0.1):
    """
    Create a mask for worked locations on the given coordinate grid
    """
    df = pd.read_excel(locations_file)
    
    # Create a mask array (False = empty, True = has worked location)
    mask = np.zeros((len(lat_coords), len(lon_coords)), dtype=bool)

    for _, row in df.iterrows():
        lat, lon = row['Latitude'], row['Longitude']
        
        # Find closest grid cell
        lat_idx = np.argmin(np.abs(lat_coords - lat))
        lon_idx = np.argmin(np.abs(lon_coords - lon))
        
        # Check if within half grid step (to ensure it's in the right cell)
        if (np.abs(lat_coords[lat_idx] - lat) <= grid_step/2 and 
            np.abs(lon_coords[lon_idx] - lon) <= grid_step/2):
            mask[lat_idx, lon_idx] = True
    
    return mask

def find_coordinate_names(ds):
    """
    Find the actual coordinate names in the dataset
    """
    lat_names = []
    lon_names = []
    
    # Check both coordinates and dimensions
    all_names = list(ds.coords.keys()) + list(ds.dims.keys())
    
    for name in all_names:
        name_lower = name.lower()
        if name_lower in ['lat', 'latitude', 'y']:
            lat_names.append(name)
        elif name_lower in ['lon', 'longitude', 'x']:
            lon_names.append(name)
    
    # Also check data variables that might be coordinates
    for var_name, var_data in ds.data_vars.items():
        if len(var_data.dims) == 1:  # 1D variables might be coordinates
            var_name_lower = var_name.lower()
            if var_name_lower in ['lat', 'latitude', 'y']:
                lat_names.append(var_name)
            elif var_name_lower in ['lon', 'longitude', 'x']:
                lon_names.append(var_name)
    
    return lat_names, lon_names

def standardize_coord_names(ds):
    """
    Standardize coordinate names to 'lat' and 'lon'
    """
    lat_names, lon_names = find_coordinate_names(ds)
    
    print(f"  Found potential lat coordinates: {lat_names}")
    print(f"  Found potential lon coordinates: {lon_names}")
    
    coord_mapping = {}
    
    # Use the first found coordinate name
    if lat_names and lat_names[0] != 'lat':
        coord_mapping[lat_names[0]] = 'lat'
    if lon_names and lon_names[0] != 'lon':
        coord_mapping[lon_names[0]] = 'lon'
    
    if coord_mapping:
        print(f"  Renaming coordinates: {coord_mapping}")
        ds = ds.rename(coord_mapping)
    
    return ds

def process_nc4_file(input_file, output_file, bbox, worked_mask, lat_coords, lon_coords, compression_level=9, planting_start_year=None):
    """
    Process a single NC4 file to create sparse array
    """
    try:
        # Try different ways to open the file
        try:
            ds = xr.open_dataset(input_file, decode_times=False)
        except:
            # Try with netcdf4 engine
            try:
                ds = xr.open_dataset(input_file, decode_times=False, engine='netcdf4')
            except:
                # Try with h5netcdf engine
                ds = xr.open_dataset(input_file, decode_times=False, engine='h5netcdf')
        
        print(f"  Dataset info:")
        print(f"    Coordinates: {list(ds.coords.keys())}")
        print(f"    Dimensions: {list(ds.dims.keys())}")
        print(f"    Data variables: {list(ds.data_vars.keys())}")
        
        # Trim time dimension if planting_start_year is provided and time coordinate exists
        if planting_start_year is not None and "time" in ds.coords:
            base = os.path.basename(input_file).lower()
            
            # Handle daily weather data (tasmax, tasmin, pr, chirps, etc.)
            if any(k in base for k in ("tasmax", "tasmin", "pr","chirps","tempmax","tempmin")):
                # time is already datetime64[ns] per your note
                # 1) Ensure sorted & unique time
                ds = ds.sortby("time")
                if hasattr(ds.get_index("time"), "duplicated"):
                    ds = ds.sel(time=~ds.get_index("time").duplicated())
                ds = xr.decode_cf(ds, use_cftime=False)  # yields numpy datetime64 when possible
                ds  = ds.sortby("time")                   # good hygiene
                # 2) Do the range in a single inclusive slice
                start_year = int(planting_start_year) - 1
                cutoff = f"{start_year}-01-01"
                end_year = 2022
                end_cutoff = f"{end_year+1}-06-30"
                before = ds.dims.get("time", 0)
                t = ds["time"]
                mask = (t >= np.datetime64(cutoff)) & (t <= np.datetime64(end_cutoff))
                ds = ds.sel(time=mask)
                after = ds.dims.get("time", 0)
                print(f"  trimmed {base}: time {before} → {after} (>= {planting_start_year}-01-01)")
            
            # Handle yearly data (planting dates, yield, etc.)
            elif any(k in base for k in ("plant", "pldate", "yield", "harvest")):
                t = ds["time"]
                units_str = str(t.attrs.get("units", "")).lower()
                
                # Check if time units are in years
                if "year" in units_str:
                    before = ds.dims.get("time", 0)
                    
                    # Parse the reference year from units like "years since 1981" or just "years"
                    ref_year_match = re.search(r"years?\s+since\s+(\d{4})", units_str)
                    if ref_year_match:
                        # "years since YYYY" format
                        ref_year = int(ref_year_match.group(1))
                        # time values are offsets from reference year
                        actual_years = ref_year + t.values.astype(int)
                    else:
                        # Assume time values are absolute years (e.g., 1981, 1982, ...)
                        actual_years = t.values.astype(int)
                    
                    # Filter to years between planting_start_year - 1 and 2022
                    start_year = int(planting_start_year) - 1
                    end_year = 2022
                    
                    year_mask = (actual_years >= start_year) & (actual_years <= end_year)
                    ds = ds.isel(time=year_mask)
                    
                    after = ds.dims.get("time", 0)
                    print(f"  trimmed {base}: time {before} → {after} years (from {start_year} to {end_year})")

        # Standardize coordinate names
        ds = standardize_coord_names(ds)
        
        # Check if we have the required coordinates after standardization
        if 'lat' not in ds.coords or 'lon' not in ds.coords:
            print(f"  Warning: Could not find or create lat/lon coordinates")
            # Try to find them in data variables and promote to coordinates
            lat_names, lon_names = find_coordinate_names(ds)
            if lat_names and lon_names:
                # Promote data variables to coordinates
                ds = ds.set_coords(lat_names + lon_names)
                ds = standardize_coord_names(ds)
            
            if 'lat' not in ds.coords or 'lon' not in ds.coords:
                return False
        
        print(f"  Using coordinates: lat={ds.lat.shape}, lon={ds.lon.shape}")
        
        # Subset to bounding box
        try:
            ds_subset = ds.sel(
                lat=slice(bbox['min_lat'], bbox['max_lat']),
                lon=slice(bbox['min_lon'], bbox['max_lon'])
            )
        except:
            # Try with indexing if slicing fails
            lat_mask = (ds.lat >= bbox['min_lat']) & (ds.lat <= bbox['max_lat'])
            lon_mask = (ds.lon >= bbox['min_lon']) & (ds.lon <= bbox['max_lon'])
            ds_subset = ds.where(lat_mask & lon_mask, drop=True)
        
        # Get the actual lat/lon values after subsetting
        actual_lats = ds_subset.lat.values
        actual_lons = ds_subset.lon.values
        
        # Handle 1D vs 2D coordinate arrays
        if actual_lats.ndim == 1 and actual_lons.ndim == 1:
            # 1D coordinates - create meshgrid for mask interpolation
            lon_grid, lat_grid = np.meshgrid(actual_lons, actual_lats)
            mask_interp = np.zeros_like(lat_grid, dtype=bool)
            
            for i in range(lat_grid.shape[0]):
                for j in range(lat_grid.shape[1]):
                    lat_val, lon_val = lat_grid[i, j], lon_grid[i, j]
                    lat_idx = np.argmin(np.abs(lat_coords - lat_val))
                    lon_idx = np.argmin(np.abs(lon_coords - lon_val))
                    mask_interp[i, j] = worked_mask[lat_idx, lon_idx]
        else:
            # 2D coordinates
            mask_interp = np.zeros_like(actual_lats, dtype=bool)
            for i in range(actual_lats.shape[0]):
                for j in range(actual_lats.shape[1]):
                    lat_val, lon_val = actual_lats[i, j], actual_lons[i, j]
                    lat_idx = np.argmin(np.abs(lat_coords - lat_val))
                    lon_idx = np.argmin(np.abs(lon_coords - lon_val))
                    mask_interp[i, j] = worked_mask[lat_idx, lon_idx]
        
        # Apply mask to all data variables
        for var_name, var_data in ds_subset.data_vars.items():
            if 'lat' in var_data.dims and 'lon' in var_data.dims:
                print(f"  Masking variable: {var_name} with shape {var_data.shape}")
                
                # Create expanded mask for variables with additional dimensions (like time)
                var_mask = mask_interp
                if var_data.ndim > 2:
                    # Find the position of lat and lon dimensions
                    lat_dim_pos = var_data.dims.index('lat')
                    lon_dim_pos = var_data.dims.index('lon')
                    
                    # Create the correct shape for broadcasting
                    new_shape = [1] * var_data.ndim
                    new_shape[lat_dim_pos] = mask_interp.shape[0]
                    new_shape[lon_dim_pos] = mask_interp.shape[1]
                    
                    var_mask = mask_interp.reshape(new_shape)
                    var_mask = np.broadcast_to(var_mask, var_data.shape)
                
                # Apply mask (set non-worked locations to NaN)
                ds_subset[var_name] = xr.where(var_mask, ds_subset[var_name], np.nan)
        
        # Prepare encoding for compression
        encoding = {}
        for var_name, var_data in ds_subset.data_vars.items():
            if var_data.dtype.kind == 'f':  # floating point
                encoding[var_name] = {
                    'zlib': True,
                    'complevel': compression_level,
                    'shuffle': True,
                    'dtype': 'float32'  # Use float32 instead of float64 to save space
                }
            else:
                encoding[var_name] = {
                    'zlib': True,
                    'complevel': compression_level,
                    'shuffle': True
                }
        
        # Also encode coordinates for compression
        for coord_name in ['lat', 'lon']:
            if coord_name in ds_subset.coords:
                encoding[coord_name] = {
                    'zlib': True,
                    'complevel': compression_level,
                    'dtype': 'float32'
                }
        
        # Save the sparse array with compression
        ds_subset.to_netcdf(output_file, encoding=encoding, format='NETCDF4')
        
        # Print file size information
        original_size = os.path.getsize(input_file) / (1024*1024)  # MB
        compressed_size = os.path.getsize(output_file) / (1024*1024)  # MB
        compression_ratio = original_size / compressed_size if compressed_size > 0 else 0
        print(f"  File size: {original_size:.1f}MB → {compressed_size:.1f}MB (compression ratio: {compression_ratio:.1f}x)")
        
        ds.close()
        ds_subset.close()
        
        return True
        
    except Exception as e:
        print(f"  Error processing {input_file}: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
import win32com.client
def resolve_shortcut(lnk_path):
    shell = win32com.client.Dispatch("WScript.Shell")
 
    shortcut = shell.CreateShortcut(lnk_path)
    return shortcut.TargetPath
def create_sparse_arrays(locations_file, input_folder, output_folder, grid_step=0.1, compression_level=9):
    """
    Main function to create sparse arrays for all NC4 files
    """
    # Create output folder
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Find minimal bounding box
    print("Finding minimal bounding box...")
    bbox = find_minimal_bounding_box(locations_file, grid_step)
    print(f"Bounding box: {bbox}")
    
    # Create coordinate arrays for the bounding box
    lat_coords = np.arange(bbox['min_lat'], bbox['max_lat'] + grid_step, grid_step)
    lon_coords = np.arange(bbox['min_lon'], bbox['max_lon'] + grid_step, grid_step)
    
    # Create worked locations mask
    print("Creating worked locations mask...")
    worked_mask = get_worked_locations_mask(locations_file, lat_coords, lon_coords, grid_step)
    print(f"Found {np.sum(worked_mask)} worked grid cells out of {worked_mask.size} total cells")
    
    # Process all NC4 files
    input_path = Path(input_folder)
    nc4_files = list(input_path.glob("*")) + list(input_path.glob("*.nc4"))
    
    print(f"Found {len(nc4_files)} NC4 files to process...")
    print(f"Using compression level: {compression_level}")
    
    successful = 0
    total_original_size = 0
    total_compressed_size = 0

    # Determine planting start year for potential trimming      
    planting_start_year = _planting_start_year_from_folder(input_folder)
    if planting_start_year is not None:
        print(f"[info] Trimming weather files to years >= {planting_start_year}")
    else:
        print("[warn] Could not determine planting start year; will not trim weather files.")

    for nc4_file in nc4_files:
        # Resolve shortcut if necessary
        if nc4_file.suffix.lower() == '.lnk':
            target_path = resolve_shortcut(str(nc4_file))
            if target_path and os.path.isfile(target_path):
                nc4_file = Path(target_path)
            else:
                print(f"  Warning: Could not resolve shortcut {nc4_file}, skipping.")
                continue
        
        output_file = Path(output_folder) / nc4_file.name
        print(f"\nProcessing: {nc4_file.name}")
        
        original_size = os.path.getsize(nc4_file) / (1024*1024)  # MB
        total_original_size += original_size
        
        if process_nc4_file(
            input_file=nc4_file,
            output_file=output_file,
            bbox=bbox,
            worked_mask=worked_mask,
            lat_coords=lat_coords,
            lon_coords=lon_coords,
            compression_level=compression_level,
            planting_start_year=planting_start_year):  # <<< NEW
                successful += 1
                compressed_size = os.path.getsize(output_file) / (1024*1024)  # MB
                total_compressed_size += compressed_size
                print(f"  ✓ Successfully created sparse array")
        else:
            print(f"  ✗ Failed to process file")
    
    # Print overall compression statistics
    overall_ratio = total_original_size / total_compressed_size if total_compressed_size > 0 else 0
    space_saved = total_original_size - total_compressed_size
    print(f"\nCompleted! Successfully processed {successful} out of {len(nc4_files)} files")
    print(f"Total original size: {total_original_size:.1f}MB")
    print(f"Total compressed size: {total_compressed_size:.1f}MB")
    print(f"Space saved: {space_saved:.1f}MB ({(space_saved/total_original_size)*100:.1f}%)")
    print(f"Overall compression ratio: {overall_ratio:.1f}x")
    print(f"Sparse arrays saved to: {output_folder}")

# Run the script
if __name__ == "__main__":
    locations_file = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\NKASI - MAIZE - revised 2025.xlsx"
    input_folder = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\Data_Maize_1981_2022"
    output_folder = r"./Data_Maize_1981_2022_SPARSE"
    # If Latitude/Longitude columns are missing, try to create them from 'GPS Coordinates' (formatted "Lat,Lon")
    df = pd.read_excel(locations_file)
    if 'Latitude' not in df.columns or 'Longitude' not in df.columns:
        if 'GPS Coordinates' in df.columns:
            coords = df['GPS Coordinates'].astype(str)
            lat_vals = pd.to_numeric(coords.str.split(',', n=1).str[0].str.strip(), errors='coerce')
            lon_vals = pd.to_numeric(coords.str.split(',', n=1).str[1].str.strip(), errors='coerce')
            if 'Latitude' not in df.columns:
                df['Latitude'] = lat_vals
            if 'Longitude' not in df.columns:
                df['Longitude'] = lon_vals
        else:
            # create empty numeric columns if nothing to parse
            if 'Latitude' not in df.columns:
                df['Latitude'] = np.nan
            if 'Longitude' not in df.columns:
                df['Longitude'] = np.nan
    df.to_excel(locations_file, index=False)
    #[Note: I haven't looked into compression. The LLM (I believe o4-mini worked the best, but it didnt comment much, so I used gemeni to add comments/readability and printouts.) says this is 'maximum compression' and it works quite well for the moment in reducing file size by ~20-40x.] Use maximum compression (level 9) for best file size reduction
    create_sparse_arrays(locations_file, input_folder, output_folder, grid_step=0.1, compression_level=9)