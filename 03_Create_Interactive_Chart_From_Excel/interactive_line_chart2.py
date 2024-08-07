## 2 line chart each for: 
## sales trend over time by country
## total sales over time by all countries

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

    # Create a line chart for sales trend over time for different countries
    fig_line_countries = px.line(
        sales_by_date_country,
        x='Date',
        y='Sales',
        color='Country',  # Different lines for each country
        title='Sales Trend Over Time by Country',
        labels={'Date': 'Date', 'Sales': 'Total Sales', 'Country': 'Country'},
        line_shape='linear'  # You can use 'linear' or 'spline' for smoothness
    )

    # Customize the layout for the country-specific chart
    fig_line_countries.update_layout(
        xaxis_title='Date',
        yaxis_title='Total Sales',
        title_x=0.5,
        legend_title_text='Country'
    )

    # Group the data by date only to calculate the total sales across all countries
    total_sales_by_date = data.groupby('Date')['Sales'].sum().reset_index()

    # Create a line chart for total sales trend over time
    fig_line_total = px.line(
        total_sales_by_date,
        x='Date',
        y='Sales',
        title='Total Sales Trend Over Time',
        labels={'Date': 'Date', 'Sales': 'Total Sales'},
        line_shape='linear'
    )

    # Customize the layout for the total sales chart
    fig_line_total.update_layout(
        xaxis_title='Date',
        yaxis_title='Total Sales',
        title_x=0.5
    )

    # Save and show the charts
    fig_line_countries.write_html("Sales_Trend_Over_Time_by_Country.html")
    fig_line_total.write_html("Total_Sales_Trend_Over_Time.html")
    fig_line_countries.show()
    fig_line_total.show()

except FileNotFoundError:
    print("Financial_Data.xlsx not found. Please check the file path and try again.")
except ValueError as e:
    print(e)

# Make sure to run this script from the correct directory if using command line
# cd "C:\Users\ofschlam\OneDrive - Aibel\Desktop\Github projects\automate-office-tasks-using-chatgpt-python\03_Create_Interactive_Chart_From_Excel"
