import pandas as pd
import numpy as np
from geopy.distance import geodesic

def calculate_distance_meters(lat1, lon1, lat2, lon2):
    """Calculate distance between two points in meters using geodesic distance"""
    return geodesic((lat1, lon1), (lat2, lon2)).meters

def match_villages_to_pixels(highlands_file, metadata_file, output_file):
    # Read the data files
    highlands_df = pd.read_excel(highlands_file)
    metadata_df = pd.read_csv(metadata_file)

    # Get all locations
    all_locations = highlands_df[['Village', 'Region', 'District', 'Latitude', 'Longitude']].dropna()

    results = []
    
    for _, location_row in all_locations.iterrows():
        village_name = location_row['Village']
        village_region = location_row['Region']
        village_district = location_row['District']
        village_lat = location_row['Latitude']
        village_lon = location_row['Longitude']
        
        # Count farmers at this location
        farmer_count = highlands_df[(highlands_df['Latitude'] == village_lat) & 
                                       (highlands_df['Longitude'] == village_lon)]["Farming_Size"]
        
        # Calculate distances to all pixels
        distances = []
        for _, pixel_row in metadata_df.iterrows():
            distance = calculate_distance_meters(
                village_lat, village_lon,
                pixel_row['Lat'], pixel_row['Lon']
            )
            distances.append(distance)
        
        # Find the pixel with minimum distance
        min_distance_idx = np.argmin(distances)
        closest_pixel = metadata_df.iloc[min_distance_idx]['pixel']
        
        results.append({
            'Farmer_Count': farmer_count,
            'Region': village_region,
            'District': village_district,
            'Village': village_name,
            'Latitude': village_lat,
            'Longitude': village_lon,
            'Pixel': closest_pixel
        })
    
    # Create output dataframe and save to CSV
    output_df = pd.DataFrame(results)
    output_df.to_csv(output_file, index=False)
    
    return output_df

# Run the matching
result = match_villages_to_pixels(
    r'.\data\Worked_Locations - Highlands Zone_Final.xlsx',
    r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\OutputCalendarDays180_MgtRiz_highfert_1982_2022_SPARSE\ThreeVariableContiguous-SyntheticYield-Optimistic-metadata.csv",
    r'.\data\village_pixel_matches.csv'
)

print(f"Matched {len(result)} unique villages to pixels")
print(result.head())