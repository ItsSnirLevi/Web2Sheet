# Web2Sheet 

Web2Sheet is a Python application designed to scrape car information from the CarWiz website (a platform for buying used cars) and organize it into an Excel file. This project utilizes web scraping techniques to extract details such as car model, engine specifications, and price from individual car listings on CarWiz.

## Features

- **URL Input:** Users can input URLs of car listings from the CarWiz website.
- **Excel Integration:** The extracted data is stored in an Excel file for easy organization and analysis.
- **File Management:** Users can choose existing Excel files or create new ones to store the scraped data.

![image](https://github.com/ItsSnirLevi/Web2Sheet/assets/127433228/97c8ec4c-7647-4e3a-a13b-89a80d9f69cd)

## Requirements

- Python 3.x
- Libraries: `requests`, `tkinter`, `selenium`, `pandas`, `webdriver_manager`

## Installation

1. Clone the repository to your local machine:  
`git clone https://github.com/your-username/web2sheet.git`

2. Install the required Python libraries using pip:  
`pip install requests tk selenium pandas webdriver_manager`

3. Ensure you have Google Chrome installed on your machine, as the project utilizes the Chrome WebDriver.

## Usage

1. Cd over to the project directory:
2. Run the web2sheet.py script:
`python program.py`
3. Choose an existing Excel file or create a new one to store the scraped data.
4. Enter the URL of a car listing from the CarWiz website.
5. Click the "Create file and Add Data" button to scrape and add the data to the Excel file.
6. Monitor the status label for feedback on the scraping process.  
**Important: Ensure that the Excel file is closed before attempting to update it with new data; otherwise, it won't work.**

![image](https://github.com/ItsSnirLevi/Web2Sheet/assets/127433228/085e88a9-30c9-429f-b519-02d5d6a0a8dd)

## Future Work

- **Optimization for Speed:** Investigate and implement methods to improve the speed of data scraping and processing to reduce execution time.
- **Enhanced Feature Collection:** Explore additional features available on the CarWiz website and integrate them into the Excel sheet, providing more comprehensive information about the listed cars.
- **Excel File Styling:** Implement styling options to enhance the visual presentation of the Excel file, making it more user-friendly and visually appealing. This could include formatting cells, adding colors, and incorporating logos or branding elements.

*_Icon created by afif fudin - Flaticon_
