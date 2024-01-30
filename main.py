import os
import requests
import pandas as pd

# Function to fetch NPI data based on first name and last name
def get_npi_data(first_name, last_name):
    url = "https://npiregistry.cms.hhs.gov/api/"
    params = {
        "first_name": first_name,
        "last_name": last_name,
        "version": "2.1"
    }

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()

        if data["result_count"] > 0:
            return data["results"][0]["number"], "Success"
        else:
            return None, "No Match"
    except requests.exceptions.RequestException as e:
        return None, f"Error: {e}"

# Load Excel file
excel_file_path = r"C:\Users\truvi\workspace\CareHigh-Python Related Projects\NPIRegisterAPIData\DataTeam\Main.xlsx"
df = pd.read_excel(excel_file_path)

# Create a new column 'NPI' and map NPI numbers based on first name and last name
df["NPI"] = ""

# Extract the base name of the loaded Excel file
base_file_name = os.path.splitext(os.path.basename(excel_file_path))[0]

# Check for the existence of an intermediate file to resume from the last saved state
intermediate_excel_path = f"C:\\Users\\truvi\\workspace\\CareHigh-Python Related Projects\\NPIRegisterAPIData\\DataTeam\\{base_file_name}_intermediate_excel_file.xlsx"
resume_info_path = f"C:\\Users\\truvi\\workspace\\CareHigh-Python Related Projects\\NPIRegisterAPIData\\DataTeam\\{base_file_name}_resume_info.txt"

resume_index = 0
if os.path.exists(resume_info_path):
    with open(resume_info_path, 'r') as file:
        resume_index = int(file.read())

for index, row in df.iloc[resume_index:].iterrows():
    npi, status = get_npi_data(row["firstName"], row["lastName"])
    df.at[index, "NPI"] = npi
    print(f"Processed record {index + 1}/{len(df)} - Status: {status} - NPI: {npi}")

    # Save the progress (current index) and the updated DataFrame to the intermediate Excel file
    with pd.ExcelWriter(intermediate_excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

    # Save the progress (current index) to a file
    with open(resume_info_path, 'w') as file:
        file.write(str(index + 1))

# Save the final dataframe to Excel
output_excel_path = f"C:\\Users\\truvi\\workspace\\CareHigh-Python Related Projects\\NPIRegisterAPIData\\DataTeam\\{base_file_name}_output_excel_file.xlsx"
df.to_excel(output_excel_path, index=False)

# Remove the intermediate file as it's no longer needed
if os.path.exists(intermediate_excel_path):
    os.remove(intermediate_excel_path)

# Remove the resume_info file as it's no longer needed
if os.path.exists(resume_info_path):
    os.remove(resume_info_path)

# Display the result
print("NPI numbers added to the Excel file:", output_excel_path)
