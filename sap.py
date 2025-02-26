import re
import pandas as pd
from tkinter import Tk, Label, Entry, Button, filedialog, StringVar, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Function to extract data from the webpage after clicking the button
def extract_data_from_web(url):
    try:
        # Set up Selenium WebDriver
        service = Service("path/to/chromedriver")  # Replace with the path to your ChromeDriver
        driver = webdriver.Chrome(service=service)

        # Open the webpage
        driver.get(url)

        # Wait for the button to be clickable and click it
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "__button54-img"))
        )
        button.click()

        # Wait for the dynamic content to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Giriş Saati')]"))
        )

        # Extract the page source after the content has loaded
        html_content = driver.page_source

        # Close the browser
        driver.quit()

        # Parse the HTML content
        soup = BeautifulSoup(html_content, 'html.parser')
        text = soup.get_text()

        # Regular expression to extract date, Giriş Saati, and Çıkış Saati
        pattern = r"(\d{1,2} \w+ \d{4}) Added externalCode \d+ Giriş Saati (\d{2}:\d{2}:\d{2}) Çıkış Saati (\d{2}:\d{2}:\d{2})"

        # Find all matches
        matches = re.findall(pattern, text)

        # Organize data into a list of dictionaries
        data = []
        for match in matches:
            date, entrance_time, exit_time = match
            data.append({
                "Date": date,
                "Entrance Time": entrance_time,
                "Exit Time": exit_time
            })

        return data

    except Exception as e:
        messagebox.showerror("Error", f"Error fetching or parsing data: {e}")
        return []

# Function to save data to an Excel file
def save_to_excel(data, output_folder):
    if not data:
        messagebox.showwarning("No Data", "No data extracted. Cannot create Excel file.")
        return

    # Create a DataFrame
    df = pd.DataFrame(data)

    # Define the output file path
    output_file = f"{output_folder}/attendance_records.xlsx"

    # Export to Excel
    df.to_excel(output_file, index=False)

    messagebox.showinfo("Success", f"Data has been successfully exported to {output_file}")

# Function to handle the "Scrape and Save" button click
def scrape_and_save():
    url = url_var.get().strip()
    if not url:
        messagebox.showwarning("Input Error", "Please enter a valid URL.")
        return

    # Select output folder
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if not output_folder:
        messagebox.showwarning("Folder Error", "No folder selected. Please try again.")
        return

    # Extract data
    data = extract_data_from_web(url)

    # Save data to Excel
    save_to_excel(data, output_folder)

# Main function to create the GUI
def main():
    global url_var

    # Create the main window
    root = Tk()
    root.title("Web Scraper and Excel Exporter")

    # URL input
    Label(root, text="Enter Website URL:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    url_var = StringVar()
    Entry(root, textvariable=url_var, width=50).grid(row=0, column=1, padx=10, pady=10)

    # Scrape and Save button
    Button(root, text="Scrape and Save", command=scrape_and_save).grid(row=1, column=0, columnspan=2, pady=10)

    # Run the GUI loop
    root.mainloop()

if __name__ == "__main__":
    main()
