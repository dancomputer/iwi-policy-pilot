import pandas as pd
import numpy as np
import xarray as xr
import os
from pathlib import Path
import warnings
import re
warnings.filterwarnings('ignore')

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

def get_precise_locations_mask(locations_file, ds):
    """
    Creates a boolean mask on the dataset's own high-resolution grid.
    A grid cell is marked True only if it's the closest one to a worked location.
    """
    df = pd.read_excel(locations_file)
    
    # Get the coordinate arrays directly from the dataset
    lat_coords = ds.lat.values
    lon_coords = ds.lon.values
    
    # Create an empty mask with the same shape as the dataset's spatial grid
    mask = np.zeros((len(lat_coords), len(lon_coords)), dtype=bool)

    # For each location in your Excel file...
    for _, row in df.iterrows():
        lat, lon = row['Latitude'], row['Longitude']
        
        # ...find the single closest grid cell in the high-resolution grid
        lat_idx = np.argmin(np.abs(lat_coords - lat))
        lon_idx = np.argmin(np.abs(lon_coords - lon))
        
        # Set only that single cell to True
        mask[lat_idx, lon_idx] = True
        
    # Print how many unique high-res cells were found
    print(f"  Identified {np.sum(mask)} unique high-resolution grid cells corresponding to worked locations.")
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

def process_nc4_file(input_file, output_file, bbox, lat_coords, lon_coords, compression_level=9, start_year=None, end_year=None):
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
        
        # Trim time dimension if start_year is provided and time coordinate exists
        if start_year is not None and "time" in ds.coords:
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
                cutoff = f"{start_year}-01-01"
                end_cutoff = f"{end_year+1}-06-30"
                before = ds.dims.get("time", 0)
                t = ds["time"]
                mask = (t >= np.datetime64(cutoff)) & (t <= np.datetime64(end_cutoff))
                ds = ds.sel(time=mask)
                after = ds.dims.get("time", 0)
                print(f"  trimmed {base}: time {before} → {after} (>= {start_year}-01-01) and <= {end_year}-06-30)")
            
            # Handle yearly data (planting dates, yield, etc.)
            if any(k in base for k in ("plantday", "pldate", "yield", "harvest")):
                print("  Detected yearly data, trimming time dimension...")
                t = ds["time"]

                # Assume time values are absolute years (e.g., 1981, 1982, ...)
                print("  Assuming time values are absolute years")
                actual_years = t.values.astype(int)
                # Filter to years between start_year and end_year
                if end_year == actual_years.max():
                    print("  end_year matches max year in data, using open-ended range")
                    year_mask = (actual_years >= start_year)
                else:
                    year_mask = (actual_years >= start_year) & (actual_years <= end_year)
                before = ds.dims.get("time", 0)
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
        
        # Apply mask to all data variables
        for var_name, var_data in ds_subset.data_vars.items():
            if 'lat' in var_data.dims and 'lon' in var_data.dims:
                var_mask = get_precise_locations_mask(locations_file, ds_subset[var_name])
                print(f"  Masking variable: {var_name} with shape {var_data.shape}")
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
def create_sparse_arrays(locations_file, input_folder, output_folder, grid_step=0.1, compression_level=9, start_year=1981, end_year=2022):
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
    #print(f"Found {np.sum(worked_mask)} worked grid cells out of {worked_mask.size} total cells")
    
    # Process all NC4 files
    input_path = Path(input_folder)
    nc4_files = list(input_path.glob("*")) # + list(input_path.glob("*.nc4"))
    
    print(f"Found {len(nc4_files)} NC4 files to process...")
    print(f"Using compression level: {compression_level}")
    
    successful = 0
    total_original_size = 0
    total_compressed_size = 0

    for nc4_file in nc4_files:
        # Resolve shortcut if necessary
        if nc4_file.suffix.lower() == '.lnk':
            target_path = resolve_shortcut(str(nc4_file))
            output_file = Path(output_folder) / (nc4_file.stem + ".nc4")
            if target_path and os.path.isfile(target_path):
                nc4_file = Path(target_path)
            else:
                print(f"  Warning: Could not resolve shortcut {nc4_file}, skipping.")
                continue
        else:
            output_file = Path(output_folder) / nc4_file.name
        
        print(f"\nProcessing: {nc4_file.name}")
        
        original_size = os.path.getsize(nc4_file) / (1024*1024)  # MB
        total_original_size += original_size
        
        if process_nc4_file(
            input_file=nc4_file,
            output_file=output_file,
            bbox=bbox,
            lat_coords=lat_coords,
            lon_coords=lon_coords,
            compression_level=compression_level,
            start_year=start_year,
            end_year=end_year):  # <<< NEW
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
    start_year = 1981 #planting start year
    end_year = 2022 #planting end year
    # Define file paths
    locations_file = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\NKASI - MAIZE - revised 2025.xlsx"
    input_folder = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\Data_Maize_1981_2023"
    output_folder = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\Data_Maize_{start_year}_{end_year}_SPARSE".format(start_year=start_year, end_year=end_year)
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
    create_sparse_arrays(locations_file, input_folder, output_folder, grid_step=0.1, compression_level=9, start_year=start_year, end_year=end_year)