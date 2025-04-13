import pandas as pd
import plotly.express as px
import streamlit as st
import os

# Function to load the Excel data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name=None)  # Read all sheets into a dictionary of DataFrames
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

# Helper function to coerce data to numeric
def coerce_to_numeric(data, column):
    try:
        data[column] = pd.to_numeric(data[column], errors='coerce')
    except Exception as e:
        st.warning(f"Could not convert {column} to numeric: {e}")
    return data

# Function to aggregate data
def aggregate_data(data, group_col, agg_col, agg_func):
    if agg_func == "Count":
        return data.groupby(group_col).size().reset_index(name='Count')
    elif agg_func == "Unique":
        unique_counts = data.groupby(group_col)[agg_col].apply(lambda x: x.astype(str).nunique()).reset_index(name=f'Unique {agg_col}')
        return unique_counts
    elif agg_func in ["Sum", "Mean", "Median", "Max", "Min"]:
        data = coerce_to_numeric(data, agg_col)
        agg_func_map = {
            "Sum": data.groupby(group_col, as_index=False)[agg_col].sum,
            "Mean": data.groupby(group_col, as_index=False)[agg_col].mean,
            "Median": data.groupby(group_col, as_index=False)[agg_col].median,
            "Max": data.groupby(group_col, as_index=False)[agg_col].max,
            "Min": data.groupby(group_col, as_index=False)[agg_col].min,
        }
        return agg_func_map[agg_func]()
    else:
        return data

# Function to perform custom calculations
def perform_custom_calculation(data, column, calc_type):
    try:
        if calc_type in ['Count', 'Unique']:
            if calc_type == 'Count':
                result = len(data[column].astype(str))
            elif calc_type == 'Unique':
                result = data[column].astype(str).nunique()
        else:
            data = coerce_to_numeric(data, column)
            calc_func_map = {
                "Sum": data[column].sum,
                "Mean": data[column].mean,
                "Median": data[column].median,
                "Max": data[column].max,
                "Min": data[column].min,
            }
            result = calc_func_map[calc_type]()
        return result
    except Exception as e:
        st.error(f"Error performing {calc_type} calculation on {column}: {e}")
        return None

# Function to generate charts
def generate_chart(data, chart_type, x_col, y_col, y_agg, color_col=None):
    try:
        # Apply aggregation if necessary
        if y_agg != 'None':
            data = aggregate_data(data, x_col, y_col, y_agg)
            y_col = y_agg if y_agg in ["Count", f"Unique {y_col}"] else y_col
        
        # Generate chart based on type
        if chart_type == "Line Chart":
            fig = px.line(data, x=x_col, y=y_col, color=color_col, title=f"{chart_type} - {x_col} vs {y_col}")
        elif chart_type == "Bar Chart":
            fig = px.bar(data, x=x_col, y=y_col, color=color_col, title=f"{chart_type} - {x_col} vs {y_col}")
        elif chart_type == "Histogram":
            fig = px.histogram(data, x=x_col, y=y_col, color=color_col, title=f"{chart_type} - {x_col} vs {y_col}")
        elif chart_type == "Scatter Plot":
            fig = px.scatter(data, x=x_col, y=y_col, color=color_col, title=f"{chart_type} - {x_col} vs {y_col}")
        elif chart_type == "Pie Chart":
            fig = px.pie(data, names=x_col, values=y_col, title=f"{chart_type} - {x_col} vs {y_col}")
        elif chart_type == "Box Plot":
            fig = px.box(data, x=x_col, y=y_col, color=color_col, title=f"{chart_type} - {x_col} vs {y_col}")
        else:
            st.error("Unsupported chart type.")
            return None
        return fig
    except Exception as e:
        st.error(f"Error generating {chart_type}: {e}")
        return None

# Streamlit app layout
def main():
    st.title("Automated Excel to Interactive Dashboard")

    # File Upload Section
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    file_path_input = st.text_input("Or enter the Excel file path", "")

    if uploaded_file is not None:
        file_path = uploaded_file
        df = load_data(file_path)
    elif file_path_input:
        if os.path.exists(file_path_input):
            file_path = file_path_input
            df = load_data(file_path)
        else:
            st.error("The file path does not exist.")
            return
    else:
        st.warning("Please upload a file or enter a file path.")
        return

    if df:
        st.success("Excel file loaded successfully!")
        sheet_names = list(df.keys())
        sheet_name = st.selectbox("Select a Sheet", sheet_names)
        data = df[sheet_name]
        st.write(f"Data from Sheet: {sheet_name}")
        st.dataframe(data)

        # --- Modify Table Section ---
        st.subheader("Modify Table")

        # Replace empty values with the string 'NULL'
        replace_null = st.checkbox("Replace empty values with NULL")
        if replace_null:
            # Replace empty strings or NaN values with 'NULL' as text
            data = data.replace(r"^\s*$", "NULL", regex=True)  # Replaces empty or whitespace cells with 'NULL'
            data = data.fillna("NULL")  # Fill NaN values with 'NULL'
            st.write("Empty values replaced with 'NULL'.")
            st.dataframe(data)

        # Option to add serial numbers (starting from 1)
        add_serial_numbers = st.checkbox("Add Serial Numbers to Data")
        if add_serial_numbers:
            data.insert(0, 'Serial Number', range(1, len(data) + 1))  # Add serial number as the first column
            st.write("Serial numbers added as the first column.")
            st.dataframe(data)

        # --- Generate Chart Section ---
        st.subheader("Generate Chart")

        # Chart Selection
        columns = data.columns.tolist()
        x_col = st.selectbox("Select the X-axis Column", columns)
        y_col = st.selectbox("Select the Y-axis Column", ["None"] + columns)

        y_agg = st.selectbox("Y-axis Aggregation", ["None", "Count", "Unique", "Sum", "Mean", "Median", "Max", "Min"])
        color_col = st.selectbox("Select Color Column (optional)", [None] + columns)

        chart_type = st.selectbox("Select Chart Type", ["Line Chart", "Bar Chart", "Histogram", "Scatter Plot", "Pie Chart", "Box Plot"])

        # Button to generate chart
        generate_chart_button = st.button("Generate Chart")
        if generate_chart_button:
            fig = generate_chart(data, chart_type, x_col, y_col, y_agg, color_col)
            if fig:
                st.plotly_chart(fig)

        # --- Perform Custom Calculations Section ---
        st.subheader("Perform Custom Calculations")

        calc_col = st.selectbox("Select Column for Calculation", columns)
        calc_type = st.selectbox("Select Calculation Type", ["Count", "Unique", "Sum", "Mean", "Median", "Max", "Min"])

        if st.button("Perform Calculation"):
            result = perform_custom_calculation(data, calc_col, calc_type)
            if result is not None:
                st.success(f"{calc_type} of {calc_col}: {result}")

if __name__ == "__main__":
    main()
