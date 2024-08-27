import requests
import openpyxl
from python_anticaptcha import AnticaptchaClient, ImageToTextTask
import concurrent.futures

# Load the Excel file and read PAN numbers
excel_file = "C:/Users/USER/Downloads/1724684937-Demo-Sheet.xlsx"
wb = openpyxl.load_workbook(excel_file)
ws = wb.active

# Anticaptcha API configuration
ANTICAPTCHA_API_KEY = "your_api_key_here"
client = AnticaptchaClient(ANTICAPTCHA_API_KEY)


# Function to solve captcha
def solve_captcha(image_url):
    print("Solving captcha...")
    image_response = requests.get(image_url)
    task = ImageToTextTask(image_response.content)
    job = client.createTask(task)
    job.join()
    captcha_text = job.get_captcha_text()
    print("Captcha solved:", captcha_text)
    return captcha_text


# Function to scrape data for a single PAN number
def scrape_pan_data(pan_number, row):
    try:
        # Convert PAN number to string for validation
        pan_number = str(pan_number)

        if pan_number.strip() == "":
            print(f"Skipping empty PAN in row {row}")
            return

        url = "https://services.gst.gov.in/services/searchtpbypan"

        # Make a request to get the captcha image and form data
        session = requests.Session()
        response = session.get(url, timeout=10)
        print("Requesting captcha image and form data...")

        # Extract captcha URL and other necessary form data from the response
        captcha_url = ""  # Extract captcha image URL from response
        print("Captcha URL:", captcha_url)
        captcha_text = solve_captcha(captcha_url)

        # Prepare form data
        data = {
            "pan": pan_number,
            "captcha": captcha_text,
            # Add other necessary fields here
        }

        # Send the request with PAN and captcha
        result_response = session.post(url, data=data, timeout=10)
        print("Sending request with PAN and captcha...")

        # Parse the result (determine success or failure)
        if "No result found" in result_response.text:
            ws.cell(row=row, column=2).value = "No result found"
            print("No result found for PAN:", pan_number)
        else:
            # Extract and enter data into the Excel sheet
            extracted_data = ""  # Parse result_response and extract relevant data
            ws.cell(row=row, column=2).value = extracted_data
            print("Data extracted for PAN:", pan_number)

    except Exception as e:
        ws.cell(row=row, column=2).value = f"Error: {str(e)}"
        print("Error occurred for PAN:", pan_number, "-", str(e))


# Use concurrent.futures to handle multiple scraping sessions simultaneously
def main():
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        for row in range(2, ws.max_row + 1):
            pan_number = ws.cell(row=row, column=3).value
            if pan_number is None or pan_number.strip() == "":
                print(f"Skipping empty PAN in row {row}")
                continue
            print("Scraping data for PAN:", pan_number)
            futures.append(executor.submit(scrape_pan_data, pan_number, row))

        concurrent.futures.wait(futures)

    # Save the Excel file after updating it
    wb.save(excel_file)
    print("Excel file saved.")


if __name__ == "__main__":
    main()
