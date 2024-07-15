import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

# Function to clear the terminal
def clear_terminal():
    os.system('cls' if os.name == 'nt' else 'clear')

# Function to get the company website using web scraping
def get_company_website(company_name, search_engines):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    max_attempts = len(search_engines)
    initial_delay = 5
    delay = initial_delay
    search_engine_index = 0

    print(f"Searching for {company_name}...")

    for attempt in range(max_attempts):
        search_url = search_engines[search_engine_index].format(company_name)
        search_engine_name = search_engines[search_engine_index].split('.')[1]
        print(f"Attempt {attempt + 1}: Using {search_engine_name.capitalize()} - {search_url}")

        try:
            response = requests.get(search_url, headers=headers)
            print(f"Response Status: {response.status_code}")
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                if "google" in search_url:
                    for g in soup.find_all('div', class_='g'):
                        link = g.find('a', href=True)
                        if link and 'http' in link['href']:
                            print(f"Website found: {link['href']}")
                            return link['href']
                elif "bing" in search_url:
                    for li in soup.find_all('li', class_='b_algo'):
                        link = li.find('a', href=True)
                        if link and 'http' in link['href']:
                            print(f"Website found: {link['href']}")
                            return link['href']
                elif "duckduckgo" in search_url:
                    for a in soup.find_all('a', class_='result__a', href=True):
                        if a and 'http' in a['href']:
                            print(f"Website found: {a['href']}")
                            return a['href']
                print("No website found in search results.")
                return '404'
            elif response.status_code == 429:
                print("Rate limit exceeded. Switching search engine...")
                search_engine_index = (search_engine_index + 1) % len(search_engines)
                time.sleep(delay)
                delay = initial_delay  # Reset delay
            else:
                response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Request failed: {e}")
            search_engine_index = (search_engine_index + 1) % len(search_engines)
            time.sleep(delay)
            delay = initial_delay  # Reset delay

    print("Max attempts reached. No website found.")
    return '404'

# Main function to read company names from Excel and write results back to Excel
def main():
    input_file = 'C:\\Users\\#########'
    output_file = 'C:\\Users\\########'
    
    search_engines = [
        "https://www.google.com/search?q=nederland+hotel+{}",
        "https://www.bing.com/search?q=nederland+hotel+{}",
        "https://duckduckgo.com/?q=nederland+{}"
    ]

    # Load the Excel file without headers
    df = pd.read_excel(input_file, header=None)

    # Add a new column for the websites
    df['Website'] = None

    start_time = time.time()

    # Process each company (assuming the first column is the company names)
    for idx, row in df.iterrows():
        clear_terminal()
        company_name = row[0]  # Accessing the first column
        print(f"\nProcessing company {idx + 1}/{len(df)}: {company_name}")
        print("-" * 60)
        website = get_company_website(company_name, search_engines)
        df.at[idx, 'Website'] = website
        elapsed_time = time.time() - start_time
        estimated_total_time = (elapsed_time / (idx + 1)) * len(df)
        estimated_remaining_time = estimated_total_time - elapsed_time
        print("-" * 60)
        print(f"Estimated time remaining: {time.strftime('%H:%M:%S', time.gmtime(estimated_remaining_time))}")
        if website != '404':
            print(f"\nWebsite for {company_name}: {website}")
            time.sleep(1)

    # Save the results to a new Excel file
    df.to_excel(output_file, index=False, header=False)
    print(f"\nResults saved to {output_file}")

if __name__ == "__main__":
    main()
