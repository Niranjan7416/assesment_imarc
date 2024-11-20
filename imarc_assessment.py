from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd
import matplotlib.pyplot as plt
import time

# Step 1: Path for ChromeDriver
chrome_service = Service("C:\Webdrivers\chromedriver.exe")

# Step 2: Initialize the WebDriver
driver = webdriver.Chrome(service=chrome_service)

#Step 3: Navigate to the URL
url = "https://www.barchart.com/futures"
driver.get(url)

#Step 4: Wait for the page to load
time.sleep(5)

#Step 5: Extract the table using xPATH
element = driver.find_element(By.XPATH, '//*[@id="main-content-column"]/div/div[3]/div[2]')
raw_data = element.text
print('raw')
print(raw_data) 

# Split raw data into lines
lines = raw_data.strip().split("\n")

# Extract column headers dynamically
columns = lines[:lines.index("Links")]

# Extract data rows dynamically
data = lines[lines.index("Links") + 1:]

# Group data into rows based on the number of columns
rows = [data[i:i + len(columns)] for i in range(0, len(data), len(columns))]

# Create DataFrame
df = pd.DataFrame(rows, columns=columns)
print('Original DataFrame')
print(df)

# Function to handle fractions, remove commas, and convert to float
def convert_to_float(value):
    value = value.replace(',', '')  # Remove commas
    if '-' in value:
        # Handle mixed fractions like "109-195"
        parts = value.split('-')
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():  # Check if it's a valid mixed fraction
            return int(parts[0]) + int(parts[1]) / 32  # Convert mixed fraction to decimal
        else:
            # Check for negative fractions like "-0-075"
            try:
                if value.startswith('-'):
                    parts = value[1:].split('-')  # Remove the negative sign and split
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        return -float(parts[0]) - float(parts[1]) / 1000  # Convert "-0-075"
                else:
                    return float(value)  # If no valid fraction, convert normally
            except ValueError:
                return float(value)  # Fallback to normal float conversion if error occurs
    return float(value)  # Return as float if no fraction

# Apply conversion to "High" and "Low" columns
df['High'] = df['High'].apply(convert_to_float)
df['Low'] = df['Low'].apply(convert_to_float)

# Add "Mean" column
df['Mean'] = (df['High'] + df['Low']) / 2

# Convert "Change" to numeric (removing commas and handling negative values)
df['Change'] = df['Change'].str.replace(',', '').apply(lambda x: convert_to_float(x))

# Find the "Contract Name" and "Last" of the row with the largest "Change"
largest_change_row = df.loc[df['Change'].idxmax()]
largest_change_contract_name = largest_change_row['Contract Name']
largest_change_last = largest_change_row['Last']

# Save data to Excel
output_file = 'Extracted_Data.xlsx'
with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, index=False, sheet_name="Raw Data")

# Display results
print('DataFrame with column Mean:')
print(df)
print("Contract with largest change:")
print(f"Contract Name: {largest_change_contract_name}")
print(f"Last: {largest_change_last}")
print(f"Data saved to: {output_file}")

# Plotting
plt.figure(figsize=(12, 6))

# Plot "High", "Low", and "Mean"
plt.plot(df['Contract Name'], df['High'], label='High', marker='o')
plt.plot(df['Contract Name'], df['Low'], label='Low', marker='o')
plt.plot(df['Contract Name'], df['Mean'], label='Mean', marker='o')

# Add labels, title, and legend
plt.xlabel('Contract Name')
plt.ylabel('Values')
plt.title('High, Low, and Mean Comparison')
plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for readability
plt.legend()

# Show grid and plot
plt.grid(True)
plt.tight_layout()
plt.show()

driver.quit()