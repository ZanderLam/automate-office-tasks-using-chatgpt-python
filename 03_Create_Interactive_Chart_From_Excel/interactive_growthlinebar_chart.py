import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

try:
    # Load the data from the "Data" sheet of the "Financial_Data.xlsx" workbook
    data = pd.read_excel("Financial_Data.xlsx", sheet_name="Data")

    # Ensure the 'Date' column is parsed as a datetime object
    data['Date'] = pd.to_datetime(data['Date'])

    # Extract year and month from the 'Date' column
    data['Year'] = data['Date'].dt.year
    data['Month'] = data['Date'].dt.month

    # Group by year and month and calculate total sales
    sales_by_year_month = data.groupby(['Year', 'Month'])['Sales'].sum().reset_index()

    # Calculate growth rate by month
    sales_by_year_month['Growth Rate'] = sales_by_year_month.groupby('Year')['Sales'].pct_change() * 100  # Calculate growth rate percentage

    # Fill NaN values in Growth Rate (e.g., the first month of each year)
    sales_by_year_month['Growth Rate'] = sales_by_year_month['Growth Rate'].fillna(0)

    # Create subplots: one for growth rate and one for total sales
    fig = make_subplots(
        rows=1, cols=1,
        specs=[[{"secondary_y": True}]],  # Enable secondary y-axis
        subplot_titles=('Sales Growth Rate and Total Sales by Month and Year')
    )

    # Add bar chart for growth rate
    for year in sales_by_year_month['Year'].unique():
        df = sales_by_year_month[sales_by_year_month['Year'] == year]
        fig.add_trace(
            go.Bar(
                x=df['Month'],
                y=df['Growth Rate'],
                name=f'Growth Rate {year}',
                text=df['Growth Rate'].map('{:.2f}%'.format),  # Format percentage
                textposition='auto',
                hovertemplate='Month: %{x}<br>Growth Rate: %{y:.2f}%<extra></extra>'
            ),
            secondary_y=False  # Primary Y-axis for growth rate
        )

    # Add line chart for total sales
    for year in sales_by_year_month['Year'].unique():
        df = sales_by_year_month[sales_by_year_month['Year'] == year]
        fig.add_trace(
            go.Scatter(
                x=df['Month'],
                y=df['Sales'],
                mode='lines+markers',
                name=f'Total Sales {year}',
                text=df['Sales'].map('${:,.0f}'.format),  # Format sales
                hovertemplate='Month: %{x}<br>Total Sales: %{y:$,.0f}<extra></extra>'
            ),
            secondary_y=True  # Secondary Y-axis for total sales
        )

    # Update layout for combined figure
    fig.update_layout(
        title='Sales Growth Rate and Total Sales by Month and Year',
        xaxis_title='Month',
        yaxis_title='Growth Rate (%)',
        yaxis2_title='Total Sales',
        xaxis=dict(tickmode='linear', title='Month'),
        title_x=0.5
    )

    # Update secondary Y-axis settings
    fig.update_yaxes(
        title_text="Growth Rate (%)",
        secondary_y=False,
        showgrid=True
    )

    fig.update_yaxes(
        title_text="Total Sales",
        secondary_y=True,
        showgrid=True
    )

    # Save and show the combined figure
    fig.write_html("Sales_Growth_Rate_And_Total_Sales_By_Month_And_Year.html")
    fig.show()

except FileNotFoundError:
    print("Financial_Data.xlsx not found. Please check the file path and try again.")
except ValueError as e:
    print(e)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
