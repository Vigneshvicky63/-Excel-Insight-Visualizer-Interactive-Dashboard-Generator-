# Python Excel Dashboard Generator
This Streamlit application enables users to upload an Excel file, select a sheet, and visualize the data using interactive charts. It also provides summary statistics (e.g., sum, mean, and median) and supports various chart types like line charts, bar charts, pie charts, scatter plots, histograms, and box plots.
## Features
- Upload Excel Files: Upload an .xlsx file or provide a file path.
- Multiple Sheet Support: Select a sheet from the uploaded Excel file.
- Data Display: View the content of the selected sheet in a tabular format.
- Interactive Charts: Generate different chart types based on user input (e.g., Line, Bar, Pie, Scatter, Histogram, Box Plot).
- Summary Statistics: Display summary metrics such as unique values, sum, mean, and median for the selected columns.
- Data Cleaning: Automatically handles missing values in the selected columns before generating charts.
## Requirements
- Python 3.x
- Required Libraries:


   - streamlit for building the interactive dashboard.
        ```
       pip install streamlit
        ```
   - pandas for data handling and processing
       ```
       pip install pandas
       ```
   - plotly for generating interactive charts
       ```
       pip install plotly
       ```
   - openpyxl for reading Excel files
       ```
       pip install openpyxl
       ```
## Installation
To run this app, follow the steps below:

After installing the libraries, you can start your Streamlit app as follows:
 ```
 streamlit run app.py
 ```
it will automatically Navigate to the browser
### 1. Clone this repository
 ```
 https://github.com/Vigneshvicky63/Python-Excel-Dashboard-Generator.git
 ```
## Removing BOM
### If you've opened a file in PyCharm that contains a BOM and you want to remove it, follow these steps:
1. Open the file in PyCharm.
2. Go to File → File Properties → File Encoding (or simply right-click on the file tab and select File Properties).
3. In the Encoding dropdown, select UTF-8 (without BOM).
4. Click Apply or OK to save the changes. PyCharm will remove the BOM (if any) from the file and save it with the specified encoding.
### If you're using VS Code follow these steps to prevent or remove BOM:
To Remove BOM from Existing Files:
1. Open your Python file or CSV file in VS Code.
2. Click on the UTF-8 text in the bottom-right corner of the editor (or any other encoding if visible).
3. Select Save with Encoding → UTF-8 (without BOM).

## Contact

If you have any questions or suggestions, feel free to reach out to:

Author: Vignesh D

Email: vigneshlakshmi149@gmail.com



