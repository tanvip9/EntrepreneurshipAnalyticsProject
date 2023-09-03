import requests
from bs4 import BeautifulSoup
import pandas as pd

# Load CIK codes from the Excel spreadsheet

df = pd.read_excel('Management_Sample.xlsx', sheet_name='Sheet1')  

# Extract CIK codes from the 'CIK' column in the spreadsheet
cik_codes = df['CIK'].tolist()

# Initialize lists to store executive information
all_names = []
all_positions = []
all_backgrounds = []

# Function to retrieve executive information for a given CIK code
def retrieve_executive_info(cik_code):
    # Define the URL for the EDGAR filings search for the specific CIK code
    search_url = f'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik_code}'

    # Send an HTTP GET request to the search URL
    response = requests.get(search_url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content of the search results page
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find the link to the 424B filings page (adjust the link text as needed)
        # You may need to inspect the page source to locate the relevant link
        filing_link = soup.find('a', text='424B')

        if filing_link:
            # Extract the href attribute of the filing link to get the actual filing page URL
            filing_url = 'https://www.sec.gov' + filing_link['href']

            # Send an HTTP GET request to the filing URL
            response = requests.get(filing_url)

            # Check if the request was successful
            if response.status_code == 200:
                # Parse the HTML content of the filing page
                soup = BeautifulSoup(response.text, 'html.parser')

                # Find the section containing executive leadership information
                # You may need to inspect the page source to locate the relevant HTML elements
                executive_section = soup.find('div', class_='executive-leadership-section')

                # Initialize lists to store executive information for this filing
                names = []
                positions = []
                backgrounds = []

                # Loop through the executive entries
                for entry in executive_section.find_all('div', class_='executive-entry'):
                    # Extract the name, position, and background
                    name = entry.find('div', class_='executive-name').text.strip()
                    position = entry.find('div', class_='executive-position').text.strip()
                    background = entry.find('div', class_='executive-background').text.strip()

                    # Append the extracted information to the respective lists
                    names.append(name)
                    positions.append(position)
                    backgrounds.append(background)

                # Append the executive information for this filing to the overall lists
                all_names.extend(names)
                all_positions.extend(positions)
                all_backgrounds.extend(backgrounds)

            else:
                print(f"Failed to retrieve data for CIK code: {cik_code}")

        else:
            print(f"No 424B filings found for CIK code: {cik_code}")

    else:
        print(f"Failed to retrieve search results for CIK code: {cik_code}")

# Iterate through the list of CIK codes and retrieve executive information
for cik_code in cik_codes:
    retrieve_executive_info(cik_code)

# Create a Pandas DataFrame to store the data for all filings
data = {
    'Name': all_names,
    'Position': all_positions,
    'Background': all_backgrounds
}
df_result = pd.DataFrame(data)

# Display the data in a table
print(df_result)
