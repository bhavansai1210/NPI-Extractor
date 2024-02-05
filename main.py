import os
import requests
import pandas as pd

# Function to fetch NPI data based on first name and last name
def get_npi_data(first_name, last_name):
    url = "https://npiregistry.cms.hhs.gov/api/"
    params = {
        "enumeration_type": "",
        "taxonomy_description": "Dentist",
        "name_purpose": "",
        "first_name": first_name,
        "use_first_name_alias": "",
        "last_name": last_name,
        "organization_name": "",
        "address_purpose": "",
        "city": "",
        "state": "",
        "postal_code": "",
        "country_code": "",
        "limit": "",
        "skip": "",
        "pretty": "",
        "version": "2.1"
    }

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()

        if data["result_count"] > 0:
            npi_number = data["results"][0]["number"]
            npi_type = data["results"][0]["enumeration_type"]
            return npi_number, npi_type, "Success"
        else:
            return None, None, "No Match"
    except requests.exceptions.RequestException as e:
        return None, None, f"Error: {e}"

# Load Excel file and display duplicate count before removal
excel_file_path = r"/content/Vermont.xlsx"
df = pd.read_excel(excel_file_path)

# Display duplicate count before removal
duplicate_count_before = df.duplicated().sum()
print(f"Duplicate Records Count Before Removal: {duplicate_count_before}")

# Remove duplicate records
df = df.drop_duplicates()

# Display duplicate count after removal
duplicate_count_after = df.duplicated().sum()
print(f"Duplicate Records Count After Removal: {duplicate_count_after}")

# Create new columns for NPI and NPI Types
df["NPI"] = ""
df["NPI_Type_Individual"] = ""
df["NPI_Type_Organization"] = ""

# Extract the base name of the loaded Excel file
base_file_name = os.path.splitext(os.path.basename(excel_file_path))[0]

# Set paths for intermediate file and resume file based on the loaded Excel file
intermediate_excel_path = os.path.join(os.path.dirname(excel_file_path), f"{base_file_name}_intermediate_excel_file.xlsx")
resume_info_path = os.path.join(os.path.dirname(excel_file_path), f"{base_file_name}_resume_info.txt")

# Check for the existence of an intermediate file to resume from the last saved state
resume_index = 0
if os.path.exists(resume_info_path):
    with open(resume_info_path, 'r') as file:
        resume_index = int(file.read())

# Inside the loop where you are processing rows, update the assignment to the "NPI" and "NPI_Type" columns
for index, row in df.iloc[resume_index:].iterrows():
    npi, npi_type, status = get_npi_data(row["firstName"], row["lastName"])
    df.at[index, "NPI"] = npi
    df.at[index, "NPI_Type_Individual"] = npi_type if npi_type == "NPI-1" else None
    df.at[index, "NPI_Type_Organization"] = npi_type if npi_type == "NPI-2" else None
    print(f"Processed record {index + 1}/{len(df)} - Status: {status} - NPI: {npi} - NPI Type: {npi_type}")

    # Save the progress (current index) to a file
    with open(resume_info_path, 'w') as file:
        file.write(str(index + 1))

# Save the final dataframe to Excel
output_excel_path = os.path.join(os.path.dirname(excel_file_path), f"{base_file_name}_output_excel_file.xlsx")
df.to_excel(output_excel_path, index=False)

# Remove the intermediate file as it's no longer needed
if os.path.exists(intermediate_excel_path):
    os.remove(intermediate_excel_path)

# Remove the resume_info file as it's no longer needed
if os.path.exists(resume_info_path):
    os.remove(resume_info_path)

# Display the result
print("Completed Work & NPI numbers added to the Excel file:", output_excel_path)
