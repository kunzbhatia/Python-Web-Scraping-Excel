Here is a README description for your project:

---

# Comprehensive Web Data Extraction and Analysis

## Overview

This project demonstrates the development of a comprehensive web data extraction and analysis tool using Python, Selenium, Excel, and VBA. The tool is designed to scrape ETF performance data from the [Morningstar Canada website](https://www.morningstar.ca/ca/report/etf/performance.aspx?t=0P0001HWR9&lang=en-CA) and export it into an Excel file. The extracted data includes the date-time, price, trend, change in dollars, and percentage change of the Avantis US Equity ETF.

## Features

- **Web Scraping**: Automated extraction of ETF performance data from the Morningstar Canada website using Python and Selenium.
- **Data Storage**: Data is stored in an Excel file, ensuring accuracy and efficiency.
- **VBA Integration**: Utilization of VBA to manipulate and analyze the data within Excel, providing comprehensive insights.

## Data Extracted

The final content of the Excel file includes the following columns:
- **Date-Time**: The date and time of the data entry.
- **Price**: The price of the ETF.
- **Trend**: The trend indicating whether the price went up or down.
- **Change $**: The change in price in dollars.
- **Change %**: The percentage change in price.

Example data:
```
Date-Time      Price    Trend   Change $  Change %
02-07-2022     $50.85   Down    0.22      0.37
03-07-2022     $50.89   Up      0.12      0.21
04-07-2022     $50.85   Down    0.34      0.13
05-07-2022     $50.88   Up      0.11      0.18
06-07-2022     $50.90   Up      0.28      0.12
07-07-2022     $50.88   Down    0.24      0.21
08-07-2022     $50.85   Down    0.19      0.23
```

## Installation

To get started with the project, follow these steps:

1. **Clone the Repository**:
    ```sh
    git clone https://github.com/kunzbhatia/Python-Web-Scraping-and-Data-Manipulation.git
    cd Python-Web-Scraping-and-Data-Manipulation
    ```

2. **Install Dependencies**:
    ```sh
    pip install -r requirements.txt
    ```

3. **Run the Script**:
    ```sh
    python scrape_etf_data.py
    ```

## Usage

1. **Web Scraping**:
    - The script uses Selenium to open the Morningstar Canada website and extract the required data.
    - The extracted data is then saved into an Excel file.

2. **Data Manipulation with VBA**:
    - A VBA macro is used to manipulate and analyze the data within the Excel file, providing additional insights and organizing the data for better readability.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any features, bug fixes, or enhancements.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
