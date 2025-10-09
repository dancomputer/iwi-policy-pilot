import pandas as pd
import numpy as np
from geopy.distance import geodesic

def calculate_distance_meters(lat1, lon1, lat2, lon2):
    """Calculate distance between two points in meters using geodesic distance (in meters)."""
    return geodesic((lat1, lon1), (lat2, lon2)).meters

def match_villages_to_pixels(highlands_file, metadata_file, output_file):
    # Read inputs
    highlands_df = pd.read_excel(highlands_file)
    metadata_df = pd.read_csv(metadata_file)

    # Validate columns
    required_highlands = {"Village", "Region", "District", "Latitude", "Longitude", "Loan Amount"}
    required_metadata = {"pixel", "Lat", "Lon"}
    missing_h = required_highlands - set(highlands_df.columns)
    missing_m = required_metadata - set(metadata_df.columns)
    if missing_h:
        raise ValueError(f"Highlands file missing required columns: {sorted(missing_h)}")
    if missing_m:
        raise ValueError(f"Metadata file missing required columns: {sorted(missing_m)}")

    # Keep rows with valid coordinates; each row == 1 farmer
    farmers = highlands_df[["Village", "Region", "District", "Latitude", "Longitude", "Loan Amount"]].dropna(subset=["Latitude", "Longitude"]).copy()
    
    # Convert Loan Amount to numeric
    farmers["Loan Amount"] = pd.to_numeric(farmers["Loan Amount"], errors="coerce").fillna(0)

    # Preload metadata for distance calc
    meta_coords = metadata_df[["Lat", "Lon"]].to_numpy()
    meta_pixels = metadata_df["pixel"].to_numpy()

    # Map each farmer to nearest pixel
    assignments = []
    for _, r in farmers.iterrows():
        v_lat, v_lon = float(r["Latitude"]), float(r["Longitude"])
        dists = [geodesic((v_lat, v_lon), (lat, lon)).meters for lat, lon in meta_coords]
        min_idx = int(np.argmin(dists))
        assignments.append({
            "Region": r["Region"],
            "District": r["District"],
            "Village": r["Village"],
            "Latitude": v_lat,
            "Longitude": v_lon,
            "Pixel": meta_pixels[min_idx],
            "Loan Amount": r["Loan Amount"],
        })

    matches_df = pd.DataFrame(assignments)

    # ---- Data check: pixel group should NOT span multiple districts
    pixel_to_districts = matches_df.groupby("Pixel")["District"].agg(lambda s: sorted(set(map(str, s))))
    bad = pixel_to_districts[pixel_to_districts.apply(lambda ds: len(ds) > 1)]
    if not bad.empty:
        print("ERROR: The following pixel groups contain villages from multiple districts:")
        for px, districts in bad.items():
            print(f"  Pixel {px}: districts = {', '.join(districts)}")
        raise ValueError("Pixel groups spanning multiple districts detected. Please review inputs.")

    # Helper to join unique strings with "-"
    def join_unique(values):
        uniq = sorted({str(v) for v in values if pd.notna(v)})
        return "_".join(uniq) if uniq else ""

    # ---- Aggregate to one row per unique Pixel
    grouped = matches_df.groupby("Pixel").agg(
        Latitude=("Latitude", "mean"),        # avg lat of farmers in pixel
        Longitude=("Longitude", "mean"),      # avg lon of farmers in pixel
        Region=("Region", join_unique),       # joined unique regions (usually one)
        District=("District", join_unique),   # joined unique districts (validated above)
        Village=("Village", join_unique),     # "-" joined unique villages
        **{"Farmer Number": ("Village", "size")},  # count of rows (farmers) in pixel
        Villages_In_Pixel=("Village", "nunique"),
        Pixel_Loan_Amount=("Loan Amount", "sum"),
    ).reset_index()

    grouped.to_csv(output_file, index=False)
    return grouped

# Example run
if __name__ == "__main__":
    result = match_villages_to_pixels(
        r'C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\NKASI - MAIZE - revised 2025.xlsx',
        r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\OutputCalendarDays180_Maize_1982_2021_SPARSE\ThreeVariableContiguous-SyntheticYield-Optimistic-metadata.csv",
        r'.\data\village_pixel_matches_maize-nkasi.csv'
    )
    print(f"Produced {len(result)} rows (one per unique Pixel).")
    print(result.head())
