import pandas as pd
import googlemaps
from tqdm import tqdm

# Load your Excel file
file_path = 'Distance File.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Clean up the DataFrame by renaming the columns for better clarity
df.columns = ['Origin Postal', 'Origin Postal Code', 'Destination Postal Code', 'City_Name', 'Province', 'Distance (km)']

# Initialize the Google Maps client with your API key
gmaps = googlemaps.Client(key='YOUR_API_KEY')

# Function to calculate road distance between two postal codes using Google Maps Directions API
def get_road_distance(origin_postal_code, destination_postal_code):
    try:
        # Make the API request
        result = gmaps.directions(origin=origin_postal_code, destination=destination_postal_code, units='metric')

        # Extract the road distance from the response
        if result:
            distance = result[0]['legs'][0]['distance']['value'] / 1000  # Convert meters to kilometers
            return distance
        else:
            print(f"No route found for {origin_postal_code} to {destination_postal_code}")
            return None
    except Exception as e:
        print(f"Error fetching road distance: {e}")
        return None

# Function to calculate distance between two postal codes
def calculate_distance(row):
    origin_postal_code = row['Origin Postal Code']
    destination_postal_code = row['Destination Postal Code']
    distance = get_road_distance(origin_postal_code, destination_postal_code)
    return distance

# Apply the distance calculation to each row with a progress bar
print("Calculating distances...")
distances = []
try:
    for _, row in tqdm(df.iterrows(), total=len(df), unit='row'):
        distance = calculate_distance(row)
        distances.append(distance)
    df['Distance (km)'] = distances
    print("Distance calculation completed.")
except Exception as e:
    print(f"Error occurred during distance calculation: {e}")

# Save the results back to an Excel file
output_file_path = 'Distance File completed.xlsx'
try:
    df.to_excel(output_file_path, index=False)
    print(f"Distances saved to {output_file_path}")
except Exception as e:
    print(f"Error occurred while saving the Excel file: {e}")