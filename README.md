# Selenium Product Addition Automation Script

This repository contains a Python script that automates the process of adding products to an e-commerce website using Selenium WebDriver. The product details are extracted from an Excel file, allowing for efficient bulk uploads.

## Features

- **Bulk Product Upload**: Automatically adds multiple products from an Excel sheet.
- **Existence Check**: Checks if a product already exists to avoid duplicates.
- **Dynamic Element Interaction**: Waits for elements to be present and clickable, ensuring reliable interactions.
- **Data-Driven**: Uses data from an Excel file to populate product details.
- **Customizable**: Easy to modify for different e-commerce platforms.

## Requirements

- Python 3.x
- Selenium
- Pandas
- Chrome WebDriver
- An Excel file (`product_data.xlsx`) with the following columns:
  - `Product_Name`
  - `Category`
  - `Series`
  - `Type`
  - `Subtype`
  - `Color_Value`
  - `Price`
  - `Is_New_Arrival` (YES/NO)
  - `Is_On_Sale` (YES/NO)
  - `Description`

## Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/your-repo.git
