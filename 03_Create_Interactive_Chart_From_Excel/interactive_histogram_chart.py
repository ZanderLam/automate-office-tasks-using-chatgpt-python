## histogram for sales distribution by country

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

    # Create a histogram for sales distribution by country
    fig_histogram = px.histogram(
        data,
        x='Sales',
        color='Country',
        title='Sales Distribution by Country',
        labels={'Sales': 'Sales Amount', 'Country': 'Country'},
        nbins=20,  # You can adjust the number of bins for granularity
        marginal='box',  # Optional: Adds a box plot on top of the histogram for additional insights
    )

    # Customize the layout
    fig_histogram.update_layout(
        xaxis_title='Sales Amount',
        yaxis_title='Frequency',
        title_x=0.5,
        legend_title_text='Country'
    )

    # Save and show the histogram
    fig_histogram.write_html("Sales_Distribution_by_Country.html")
    fig_histogram.show()

except FileNotFoundError:
    print("Financial_Data.xlsx not found. Please check the file path and try again.")
except ValueError as e:
    print(e)

# Make sure to run this script from the correct directory if using command line
# cd "C:\Users\ofschlam\OneDrive - Aibel\Desktop\Github projects\automate-office-tasks-using-chatgpt-python\03_Create_Interactive_Chart_From_Excel"
