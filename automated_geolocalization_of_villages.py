import pandas as pd
import time
from geopy.geocoders import Nominatim

# --- 1. SETUP ---
# The full path to your input Excel file
input_excel_path = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\Worked_Locations - Highlands Zone_Final.xlsx"

# The name for the new output file, now with a .csv extension
output_csv_path = r"C:\Users\danie\NecessaryM1InternshipCode\ProjectRice\PolicyPilot\iwi-policy-pilot\data\Worked_Locations_with_Coordinates.csv"

# Initialize the geocoder with a unique user_agent
geolocator = Nominatim(user_agent="village_geocoder_app_danie")

# --- 2. READ AND PREPARE DATA ---
print(f"Reading data from '{input_excel_path}'...")
try:
    df = pd.read_excel(input_excel_path)
except FileNotFoundError:
    print(f"ERROR: The file was not found at the specified path. Please check the path and try again.")
    exit()

unique_locations = df[['Region', 'District', 'Village']].drop_duplicates().reset_index(drop=True)
print(f"Found {len(df)} total rows and {len(unique_locations)} unique locations to geocode.")


# --- 3. GEOCODE THE UNIQUE LOCATIONS ---
# Dictionary now stores a tuple of (lat, lon, status)
coordinates_cache = {}

print("\nStarting the geocoding process...")

for index, row in unique_locations.iterrows():
    query = f"{row['Village']}, {row['District']}, {row['Region']}, Tanzania"
    print(f"Processing {index + 1}/{len(unique_locations)}: {query}")
    
    location_key = (row['Region'], row['District'], row['Village'])
    
    try:
        location = geolocator.geocode(query)
        
        if location:
            # If found, store lat, lon, and "Found" status
            coordinates_cache[location_key] = (location.latitude, location.longitude, "Found")
        else:
            # If not found, store None and "Not Found" status
            coordinates_cache[location_key] = (None, None, "Not Found")
            
    except Exception as e:
        print(f"  -> An error occurred for query '{query}': {e}")
        coordinates_cache[location_key] = (None, None, f"Error: {e}")

    time.sleep(1)

print("\nGeocoding complete.")


# --- 4. MAP RESULTS BACK TO THE ORIGINAL DATAFRAME ---
print("Mapping coordinates and status back to the original data...")

# Helper function to get data from the cache
def get_data_from_cache(row, item_index):
    key = (row['Region'], row['District'], row['Village'])
    # Default to a tuple of Nones if key is somehow missing
    return coordinates_cache.get(key, (None, None, "Cache Key Not Found"))[item_index]

# Create the new columns by applying the function
df['Latitude'] = df.apply(lambda row: get_data_from_cache(row, 0), axis=1)
df['Longitude'] = df.apply(lambda row: get_data_from_cache(row, 1), axis=1)
df['Geocoding_Status'] = df.apply(lambda row: get_data_from_cache(row, 2), axis=1)


# --- 5. SAVE THE FINAL RESULTS ---
print(f"Saving the results to '{output_csv_path}'...")
df.to_csv(output_csv_path, index=False)

print("\nProcess finished successfully!")
print(f"The new CSV file with coordinates has been saved.")


# --- 6. DISPLAY SUMMARY OF NOT FOUND LOCATIONS ---
not_found_locations = []
for location_key, (lat, lon, status) in coordinates_cache.items():
    if status != "Found":
        not_found_locations.append((location_key, status))

if not_found_locations:
    print("\n--- Summary of Locations Not Found ---")
    for (region, district, village), status in not_found_locations:
        print(f"- Region: {region}, District: {district}, Village: {village} (Status: {status})")
else:
    print("\n--- All locations were successfully found! ---")