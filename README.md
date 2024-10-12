# EarningsDateUpdater
EarningsDateUpdater is a Python tool designed to automate the retrieval of stock symbols from an Excel sheet, scrape earnings dates from a financial website, and update the sheet with the corresponding information. This project showcases proficiency in web scraping with Selenium, Excel file manipulation in Python using openpyxl, and the automation of repetitive data extraction workflows.

By reducing manual effort, this tool improves efficiency, accuracy, and scalability, making it a valuable asset for financial analysts or professionals working with stock earnings data.

## Project Purpose
This project was originally developed as the final assignment for a Python course to demonstrate proficiency in web scraping, regular expressions, and Excel file manipulation, while solving a real-world problem faced by a colleague. Initially, I planned to use the `requests`, `lxml`, and `BeautifulSoup` libraries taught in the course, but due to the presence of JavaScript on the target website, I pivoted and independently learned Selenium to address this challenge.

The tool automates the extraction of stock earnings dates, a task that previously required hours of manual effort. By streamlining this workflow, the project showcases my ability to create practical, time-saving solutions using Python and highlights my adaptability in learning new technologies to overcome unexpected obstacles.

## Technologies Used
  * **Programming language:** Python
  * **Libraries:**
      * `logging`: Tracks errors and ensures the program functions as intended.
      * `openpyxl`: Facilitates interaction with Excel documents.
      * `Selenium`: Automates browser interactions and scrapes relevant data.
      * `re`: Utilizes regular expressions to locate dates.
      * `datetime`: Handles date objects.
  * **Dependencies:**
      * `Selenium WebDriver`: Requires the appropriate WebDriver for the browser used (e.g., GeckoDriver for Firefox).
  * **Version Information:**
      * Python: 3.11.5
      * `Selenium`: 4.31.1
      * `openpyxl`: 3.0.10
    
## Features
  * Utilizes Selenium best practices, including headless browsing and techniques to minimize detection.
  * Scrapes earnings dates for stocks listed in an Excel file from Zacks.com.
  * Extracts dates from two potential locations on the webpage and selects the most relevant one based on predefined criteria if both are present.
  * Inserts the selected date into the Excel file and applies custom formatting to highlight changes for easy review.

## Installation
**Requires:**
  * Python
  * `openpyxl`
  * `Selenium`
  * `GeckoDriver`

**Instructions:**
1. Download the `.py` file.
2. Place the `.py` file in the directory containing the Excel file you want to interact with.
3. Open the file in IDLE or any other IDE and modify the lines as specified in the **Usage** section.
4. Run the script in one of the following ways:
     * In IDLE or another IDE
     * Via the terminal `python earnings_date_updater.py`
     * By double-clicking the `.py` file

## Usage

  1. **Prepare an Excel sheet with stock symbols:** Ensure that the following items are configured correctly:
       * Line 23: Set the filename to match the name of the Excel file containing the stock symbols.
       * Line 34: Update `stockcella` to the column from which the first stock symbol will be read.
       * Line 35: Update `datecella` to the column where the earnings date for `stockcella` will be inserted.
       * Line 37: Update `stockcellb` to the column from which the second stock symbol will be read.
       * Line 38: Update `datecellb` to the column where the earnings date for `stockcellb` will be inserted.
       * Line 44: Set the filename to the desired name for the updated file.
  
  2. Run the script using your chosen method.

   **The script will:**
       * Extract stock symbols from the provided Excel file.
       * Scrape earnings dates from Zacks.com.
       * Update the Excel file with the earnings dates in the specified columns.
       
  3. **Review the output file:**
       * The updated file will be saved in the same directory under the specified name. Dates will be marked in purple, and any errors will be highlighted with a red fill. To prevent overwriting on future runs, save the reviewed file under a different name.

## Code Overview
  **Input:** Accepts an Excel file (`input_stocks.xlsx`) with stock symbols in columns A and B.

  **Process:** For each stock symbol, the program scrapes up to two potential earnings dates (top and bottom) from Zacks.com using Selenium. The top date is selected unless the bottom date is within the last 30 days or the top date is unavailable. If neither date is found, it returns 'MANUAL', indicating that the user must locate the date manually.

  **Output:** The Excel file is updated with the selected earnings dates in columns C and D, with the text formatted in purple. If a date is missing, an appropriate error message is inputted, and the cell is filled with red.

  **Error handling:** In the event of an error, the output cell will be filled with red. If a date cannot be located, it shows 'MANUAL'; if the webpage or its elements fail to load, it displays 'LOAD ERROR'.

## Key functions
  **`main` Function:** The control center of the program. Loads the Excel sheet, iterates through each row, and processes stock symbols. After inputting all dates, it saves the file as `output_stocks.xlsx`.

  **`search_stock` Function:** Extracts the stock symbol and skips processing if a date is already entered. If not, it calls the next function to retrieve the date and enters it into the Excel sheet.

  **`stock_process` Function:** Calls the `scrape` function for the current stock and then invokes the `select_date` function once both dates are located.

  **`scrape` Function:** Sets the WebDriver using the `set_webdriver` function, loads the page, and locates the two dates. It returns 'LOAD ERROR' if the page fails to load and closes the driver after all processes are complete.

  **`set_webdriver` Function:** Configures the WebDriver for webpage access, running in headless mode while employing best practices to minimize detection. Uses GeckoDriver for service settings and waits 10 seconds for all page elements to load.

  **`get_date` Function:** Extracts the section of the page containing the earnings date.

  **`extract_date` Function:** Uses regular expressions to isolate the date from other extracted content.

  **`select_date` Function:** Applies logic to determine which of the two dates will be entered into the Excel sheet.

  **`convert_date` Function:** Converts the date, initially in string format from the webpage, into a `datetime` date object.

  **`last_30` Function:** Determines whether the date falls within the last 30 days.

  **`apply_style` Function:** Applies formatting to the Excel sheet based on input, enhancing the visibility of changes for users.

## Error Handling
  This project employs logging to facilitate debugging by capturing detailed error information and tracking the program's execution.
  * If no earnings dates are found for a stock, the corresponding Excel cell is marked with 'MANUAL,' indicating that manual intervention is required to find the date.
  * If a page or element fails to load during the scraping process, the cell is marked with 'LOAD ERROR' and filled with red to clearly signal an issue.
  
## Relevance to Software Development:
  This project demonstrates my ability to:
  * **Automate workflows** using Python and Selenium for practical applications.
  * **Utilize webscraping and Excel manipulation** to streamline data processing and analysis.
  * **Apply robust error handling and logging**, supported by thorough testing, to address edge cases and ensure accuracy and reliability.
  * **Work with regular expressions** to parse and extract complex data patterns.
  * **Independently learn and apply new skills** outside of formal instruction (e.g., Selenium for browser automation).
  * **Design and implement decision-making logic** based on dynamic conditions (e.g., selecting the most relevant date).
  * **Write clean, modular, and maintainable code** by organizing functionality into reusable functions and adhering to best practices.
  * **Ensure scalability**, allowing the project to handle larger datasets or adapt to new use cases with minimal modifications.
  * **Manage external dependencies** (e.g., Selenium WebDriver) and configure browser settings to mitigate detection risks.
  * **Implement basic security practices** during web scraping, including user agent spoofing and anti-detection techniques.
