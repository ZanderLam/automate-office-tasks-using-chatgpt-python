#### line chart for sales trend over time including total and by country

import pandas as pd
import plotly.express as px

try:
    # Load the data from the "Data" sheet of the "Financial_Data.xlsx" workbook
    data = pd.read_excel("Financial_Data.xlsx", sheet_name="Data")

    # Check if necessary columns exist
    required_columns = ['Date', 'Sales', 'Country']
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        raise ValueError(f"Missing columns: {', '.join(missing_columns)}")

    # Convert 'Date' column to datetime format
    data['Date'] = pd.to_datetime(data['Date'])

    # Group the data by date and country, and calculate the total sales for each group
    sales_by_date_country = data.groupby(['Date', 'Country'])['Sales'].sum().reset_index()

    # Calculate total sales by date across all countries
    total_sales_by_date = data.groupby('Date')['Sales'].sum().reset_index()
    total_sales_by_date['Country'] = 'Total'  # Add a new category for total sales

    # Combine both datasets
    combined_data = pd.concat([sales_by_date_country, total_sales_by_date])

    # Create a line chart for sales trend over time including total and by country
    fig_line_combined = px.line(
        combined_data,
        x='Date',
        y='Sales',
        color='Country',  # Different lines for each country, including total
        title='Sales Trend Over Time (Including Total)',
        labels={'Date': 'Date', 'Sales': 'Total Sales', 'Country': 'Country'},
        line_shape='linear'  # You can use 'linear' or 'spline' for smoothness
    )

    # Customize the layout
    fig_line_combined.update_layout(
        xaxis_title='Date',
        yaxis_title='Total Sales',
        title_x=0.5,
        legend_title_text='Country/Total'
    )

    # Save and show the chart
    fig_line_combined.write_html("Sales_Trend_Over_Time_Combined.html")
    fig_line_combined.show()

except FileNotFoundError:
    print("Financial_Data.xlsx not found. Please check the file path and try again.")
except ValueError as e:
    print(e)

# Make sure to run this script from the correct directory if using command line
# cd "C:\Users\ofschlam\OneDrive - Aibel\Desktop\Github projects\automate-office-tasks-using-chatgpt-python\03_Create_Interactive_Chart_From_Excel"
