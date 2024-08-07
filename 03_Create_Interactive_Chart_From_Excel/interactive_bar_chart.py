import pandas as pd
import plotly.express as px
import os

try:
    # Load the data from the "Data" sheet of the "Financial_Data.xlsx" workbook
    data = pd.read_excel("Financial_Data.xlsx", sheet_name="Data")

    # Check if 'Country' and 'Sales' columns exist in data
    if 'Country' not in data.columns or 'Sales' not in data.columns:
        raise ValueError("Columns are missing")

    # Group the data by country and calculate the total sales for each country
    sales_by_country = data.groupby("Country")["Sales"].sum().reset_index()

    # Sort the data from highest to lowest sales
    sales_by_country = sales_by_country.sort_values(by="Sales", ascending=False)

    # Create the bar chart with green shades
    fig = px.bar(sales_by_country, x="Country", y="Sales", title="Financial Data By Country",
                 labels={"Country": "Country", "Sales": "Total Sales"},
                 color="Sales",  # Apply color scale based on sales
                 color_continuous_scale=px.colors.sequential.Greens)  # Use shades of green

    # Customize the layout for better presentation
    fig.update_layout(
        xaxis_title="Country",
        yaxis_title="Total Sales",
        coloraxis_colorbar=dict(title="Sales Amount"),
        title_x=0.5,  # Center the title
    )

    # Save the chart to the same directory as the workbook
    fig.write_html("Financial_Data_By_Country.html")

    # Display the chart
    fig.show()

except FileNotFoundError as e:
    print(f"{e} not found")
except ValueError as e:
    print(e)

# cd "C:\Users\ofschlam\OneDrive - Aibel\Desktop\Github projects\automate-office-tasks-using-chatgpt-python\03_Create_Interactive_Chart_From_Excel"