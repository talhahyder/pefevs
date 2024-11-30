import pandas as pd
import folium
from folium.plugins import MarkerCluster

# Load the data from the Excel file
file_path = 'data.xlsx'

# Load primary dataset (Sheet 1)
data = pd.read_excel(file_path)

# Add pie charts
def generate_pie_chart_html(male, female):
    total = male + female
    if total == 0:
        male_percentage = female_percentage = 50  # Default split if no data
    else:
        male_percentage = (male / total) * 100
        female_percentage = 100 - male_percentage

    male_arc = male_percentage * 3.6  # Degrees for male slice
    return f"""
    <div style="text-align: center;">
        <svg width="100" height="100" viewBox="0 0 32 32" style="display: block; margin: 0 auto;">
            <circle r="16" cx="16" cy="16" fill="white"></circle>
            <circle r="16" cx="16" cy="16" fill="blue"
                    stroke="blue" stroke-width="32"
                    stroke-dasharray="{male_arc} {360 - male_arc}"
                    transform="rotate(-90) translate(-32)"></circle>
            <circle r="16" cx="16" cy="16" fill="orange"
                    stroke="orange" stroke-width="32"
                    stroke-dasharray="{female_percentage * 3.6} {360 - (female_percentage * 3.6)}"
                    transform="rotate(-90) translate(-32)"></circle>
        </svg>
        <p style="margin: 5px 0; font-family: Arial; font-size: 12px;">
            <b>Male:</b> {male:,} ({male_percentage:.1f}%)<br>
            <b>Female:</b> {female:,} ({female_percentage:.1f}%)
        </p>
    </div>
    """

# Ensure required columns exist
required_columns = ['Tehsil', 'Latitude', 'Longitude', 'District', 'Province', 
                    'Total OOSC', 'Male', 'Female']
if not set(required_columns).issubset(data.columns):
    raise ValueError(f"Missing one or more required columns: {required_columns}")

# Filter rows with missing coordinates
data = data.dropna(subset=['Latitude', 'Longitude'])

# Load the 'evs' worksheet for school plotting
evs_data = pd.read_excel(file_path, sheet_name='evs')

# Load the 'Count' worksheet
count_data = pd.read_excel(file_path, sheet_name='Count')

# Create a mapping of districts to their total school counts (Column 1: District, Column 2: Total Schools)
district_school_count = dict(zip(count_data.iloc[:, 0], count_data.iloc[:, 1]))

# Ensure required columns in the 'evs' worksheet
required_evs_columns = ['Tehsil', 'District', 'School Name', 'Latitude', 'Longitude']
if not set(required_evs_columns).issubset(evs_data.columns):
    raise ValueError(f"Missing one or more required columns in 'evs' worksheet: {required_evs_columns}")

# Filter rows with missing coordinates in 'evs'
evs_data = evs_data.dropna(subset=['Latitude', 'Longitude'])

# Create a Folium map centered on the dataset's average location
m = folium.Map(location=[data['Latitude'].mean(), data['Longitude'].mean()], zoom_start=6)

# Add a heading to the map
heading_html = """
<h3 style="text-align: center; font-family: Tahoma; font-size: 14px; font-weight: bold; color: darkgrey;">
    Visualization of Out of School Children and Schools in Pakistan
</h3>
"""
m.get_root().html.add_child(folium.Element(heading_html))

# Add the source at the bottom right corner
source_html = """
<div style="
    position: fixed;
    bottom: 10px;
    right: 10px;
    background-color: rgba(255, 255, 255, 0.8);
    padding: 5px;
    border-radius: 5px;
    font-family: Tahoma;
    font-size: 12px;
    font-weight: bold;
    color: darkgrey;
    box-shadow: 0px 0px 5px rgba(0, 0, 0, 0.3);
    z-index: 10000;">
    <a href="https://mathsandscience.pk/publications/the-missing-third/" 
       style="text-decoration: none; color: darkgrey;" 
       target="_blank">
        Pak Alliance for Maths and Science
    </a>
</div>
"""
m.get_root().html.add_child(folium.Element(source_html))

# Create marker clusters for districts
primary_clusters = {}

# Add marker clusters for the primary dataset
for dist in data['District'].unique():
    primary_clusters[dist] = MarkerCluster(name=f"Primary Data - {dist}").add_to(m)

# Add markers for out-of-school data
for lat, lon, tehsil_name, children, male_count, female_count, dist in zip(
        data['Latitude'], data['Longitude'], data['Tehsil'],
        data['Total OOSC'], data['Male'], data['Female'], data['District']):
    if pd.notna(lat) and pd.notna(lon):
        # Generate pie chart HTML
        pie_chart_html = generate_pie_chart_html(male_count, female_count)

        popup_info = f"""
        <div style="max-width: 300px;">
            <b>District:</b> {dist} <br>
            <b>Tehsil:</b> {tehsil_name} <br>
            <b>Total OOSC:</b> {children:,} <br>
            {pie_chart_html}
        </div>
        """
        
        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_info, max_width=300),
            tooltip=f"Tehsil: {tehsil_name}" if pd.notna(tehsil_name) else "No Tehsil Name",
            icon=folium.Icon(color="blue", icon="info-sign"),
        ).add_to(primary_clusters[dist])

# Create a MarkerCluster for schools
school_cluster = MarkerCluster(name="Schools Cluster").add_to(m)

# Add markers for schools in the 'evs' worksheet
for lat, lon, school_name, dist, tehsil_name in zip(
        evs_data['Latitude'], evs_data['Longitude'], evs_data['School Name'],
        evs_data['District'], evs_data['Tehsil']):
    if pd.notna(lat) and pd.notna(lon):
        school_popup = f"""
        <b>School Name:</b> {school_name} <br>
        <b>District:</b> {dist} <br>
        <b>Tehsil:</b> {tehsil_name}
        """
        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(school_popup, max_width=300),
            icon=folium.Icon(color="yellow", icon="info-sign"),  # Yellow markers
        ).add_to(school_cluster)  # Add markers to the cluster

# Add layer control to toggle between clusters
folium.LayerControl().add_to(m)

# Save the map
m.save('out_of_school_map_with_schools.html')
print("Map created and saved as 'out_of_school_map_with_schools.html'")
