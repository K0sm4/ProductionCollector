
# Web Scraping and Data Extraction to Excel for Trade Items

This Python code extracts product data from two different websites using web scraping techniques and saves the extracted data into an Excel file. The code retrieves product details such as manufacturer, model, manufacturer number, and description from the websites.

The code is designed to work with two specific websites: tim.pl and tme.pl. It uses the requests and bs4 libraries to make HTTP requests and parse the HTML response, respectively. It also uses the openpyxl library to work with Excel files.
![Configuration](https://i.imgur.com/SUHZN18.png)

# How 2 use the code:
1. Clone the repository or download the main.py file.
2. Make sure you have Python 3 installed on your system.
3. Install the required libraries by running pip install -r requirements.txt in the command line.
4. Open the listofcommercialitems.xlsx file and make sure it has the following columns in the first row: Item number, Manufacturer, Model, Manufacturer number, and Description.
5. Run the code by running python main.py in the command line.
6. Enter the URL of the product page when prompted.
7. The code will extract the product data and add a new row to the Excel file with the extracted data.


# Note
The code assumes that the listofcommercialitems.xlsx file is located in the same directory as the main.py file. If it's located elsewhere, you'll need to modify the file path in the code accordingly.

The code only works with tim.pl and tme.pl. If you want to use it with other websites, you'll need to modify the code to match their HTML structure.






## Installation

To use this script, follow the steps below:

1. Clone or download the repository to your local machine.

2. Install the required Python libraries (bs4, requests, and openpyxl) using pip. You can use the following command:
```bash
  pip install bs4 requests openpyxl

```
3. Open the terminal and navigate to the directory where you have cloned or downloaded the repository.

4. Run the script using the following command:
```bash
    python3 main.py
```
5. Enter the URL of the product page when prompted.

6. The script will extract the necessary data and save it in the "listofcommercialitems.xlsx" file in the same directory.
# Contributing
If you find any issues with the code or want to suggest an improvement, feel free to submit an issue or a pull request.
