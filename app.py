from collections import namedtuple
import altair as alt
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import datetime as dt
from datetime import datetime
import io
import time
import xlsxwriter
from fpdf import FPDF
import base64



# Define your discrete color sequence PETRONAS COLORS
color_discrete_sequence = [
    "#00b1a9",  # Original color - R000 G177 B169
    "#763f98",  # Original color - R118 G063 B152
    "#20419a",  # Original color - R032 G065 B154
    "#fdb924",  # Original color - R253 G185 B036
    "#bfd730",  # Original color - R191 G215 B048
    "#007b73",  # Shade of R000 G177 B169
    "#3a1d4c",  # Shade of R118 G063 B152
    "#101e4a",  # Shade of R032 G065 B154
    "#cc8b1c",  # Shade of R253 G185 B036
    "#8e9c1b"   # Shade of R191 G215 B048
]

# Setting Up
st.set_page_config(page_title = "DashBoard",page_icon = r'Resources/4953098.png',layout ="wide")

st.markdown(
    """
        <style>
            .appview-container .main .block-container {{
                padding-top: {padding_top}rem;
                padding-bottom: {padding_bottom}rem;
                }}

        </style>""".format(
        padding_top=1, padding_bottom=1
    ),
    unsafe_allow_html=True,
)


df = pd.read_excel("PCHP Data.xlsx","Overall_data")


# Sidebar header and widgets for selecting filters
st.sidebar.header("Choose your filter:")
all_counterparties = df["FO.CounterpartyName"].dropna().unique()
all_portfolios = df["Portfolio"].dropna().unique()
all_dates = pd.to_datetime(df['FO.TradeDate']).dt.date.dropna().unique()
all_dealers = df["FO.DealerID"].dropna().unique()

# Add "All" option to the lists
all_counterparties = ['All'] + all_counterparties.tolist()
all_portfolios = ['All'] + all_portfolios.tolist()
all_dealers = ['All'] + all_dealers.tolist()

# Set default selections to include "All"
selected_counterparties = st.sidebar.multiselect("Counterparty", all_counterparties, default=['All'])
selected_portfolios = st.sidebar.multiselect("Portfolio", all_portfolios, default=['FY2024 PCHP'])

selected_dealers = st.sidebar.multiselect("Dealer", all_dealers, default=['All'])

# Update the selected options if "All" is selected
if 'All' in selected_counterparties:
    selected_counterparties = all_counterparties[1:]  # Exclude "All"
else:
    selected_counterparties = selected_counterparties

if 'All' in selected_portfolios:
    selected_portfolios = all_portfolios[1:]  # Exclude "All"
else:
    selected_portfolios = selected_portfolios

if 'All' in selected_dealers:
    selected_dealers = all_dealers[1:]  # Exclude "All"
else:
    selected_dealers = selected_dealers

# Filter data based on selected counterparties, portfolios, and date range
filtered_df = df[(df['FO.CounterpartyName'].isin(selected_counterparties)) &
                  (df['Portfolio'].isin(selected_portfolios))]

filtered_df = filtered_df[(filtered_df['FO.DealerID'].isin(selected_dealers))] 
                  

# Convert the "FO.TradeDate" column to datetime if it's not already
filtered_df['FO.TradeDate'] = pd.to_datetime(filtered_df['FO.TradeDate'], errors='coerce')



# Date range selection
st.sidebar.header("Select Date Range")

 ## Range selector
format = 'MMM DD, YYYY'  # format output

# Handle NaTType error and set default values for the date range
try:
    MIN_MAX_RANGE = (filtered_df['FO.TradeDate'].dropna().min(), filtered_df['FO.TradeDate'].dropna().max())
except KeyError:
    # Handle KeyError (e.g., due to NaTType) by setting default min and max dates
    MIN_MAX_RANGE = (pd.Timestamp('1900-01-01'), pd.Timestamp('2100-12-31'))

# Get the minimum and maximum dates from the filtered DataFrame
min_date = MIN_MAX_RANGE[0]
max_date = MIN_MAX_RANGE[1]

# Set the pre-selected dates to match the minimum and maximum dates
PRE_SELECTED_DATES = (min_date.to_pydatetime(), max_date.to_pydatetime())  # Convert to datetime objects

# Handle the KeyError (NaTType) when creating the slider
try:
    selected_min, selected_max = st.sidebar.slider(
        "Datetime slider",
        value=PRE_SELECTED_DATES,
        min_value=MIN_MAX_RANGE[0],
        max_value=MIN_MAX_RANGE[1],format=format
    )
except KeyError:
    # Set default values for the slider in case of NaTType error
    selected_min, selected_max = PRE_SELECTED_DATES

# Convert the date range to pandas Timestamp objects
start_date = pd.to_datetime(selected_min)
end_date = pd.to_datetime(selected_max)

filtered_df = filtered_df[(filtered_df['FO.TradeDate'] >= start_date) &
                  (filtered_df['FO.TradeDate'] <= end_date)]



#Title

st.title("Group Commodity Exposure Management Dashboard")
tab1, tab2, tab3, tab4 = st.tabs(["Overall Data", "Overview", "MTM", "Report"])


with tab1:
    # Display PCHP Data
    st.title("PCHP Execution Data")

    # Create a formatted copy of the filtered DataFrame to preserve the original data
    formatted_df = filtered_df.copy()

    # Format date columns for better readability
    date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']
    for column in date_columns:
        formatted_df[column] = formatted_df[column].dt.strftime('%d %b %Y')

    # Specify columns to display in the table
    columns_to_display = ['FO.TradeDate','FO.DealerID', 'FO.CounterpartyName','FO.NetPremium', 'FO.Position_Quantity',
                        'FO.StrikePrice1', 'FO.StrikePrice2', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate',
                        'E.January','E.February','E.March','E.April','E.May','E.June','E.July',
                        'E.August','E.September','E.November','E.December']
    
    

    # Reset index to start from 1
    formatted_df = formatted_df.reset_index(drop=True)

    # Start index from 1
    formatted_df.index = formatted_df.index + 1

    # Show the formatted DataFrame using st.dataframe
    st.dataframe(formatted_df[columns_to_display],height=500, use_container_width = True)

with tab2:
    st.title('Execution Overview')
    #Brent Price Data
    df_prices = pd.read_excel("PCHP Data.xlsx","Brent_Prices")
    df_prices['Date'] = pd.to_datetime(df_prices['Date'], errors='coerce')
    filtered_df_prices = df_prices[(df_prices['Date'] >= start_date) &
                    (df_prices['Date'] <= end_date)]

    # Create a Plotly line chart
    fig_Brent = px.line(filtered_df_prices, x='Date', y='Historical Brent Price', title='Trade Execution Window',
                        labels={'Historical Brent Price': 'Brent Price'})

    # Update the line color
    fig_Brent.update_traces(line_color='#808080')

    # Define a color mapping for each portfolio
    portfolio_colors = {portfolio: px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
                        for i, portfolio in enumerate(filtered_df['Portfolio'].unique())}

    # Create a dictionary to store traces for each portfolio
    portfolio_traces = {}

    # Add markers for executed trades, grouped by portfolio
    for index, trade in filtered_df.iterrows():
        portfolio = trade['Portfolio']
        if portfolio not in portfolio_traces:
            portfolio_traces[portfolio] = {'x': [], 'y': [], 'color': portfolio_colors.get(portfolio, 'red')}
        
        corresponding_price = filtered_df_prices.loc[filtered_df_prices['Date'] == trade['FO.TradeDate'], 'Historical Brent Price'].values
        if len(corresponding_price) > 0:
            portfolio_traces[portfolio]['x'].append(trade['FO.TradeDate'])
            portfolio_traces[portfolio]['y'].append(corresponding_price[0])

    # Add the portfolio traces to the figure
    for portfolio, trace_data in portfolio_traces.items():
        fig_Brent.add_trace(
            go.Scatter(
                x=trace_data['x'],
                y=trace_data['y'],
                mode='markers',
                marker=dict(color=trace_data['color']),
                name=f'Executed Trades - {portfolio}',
                legendgroup=portfolio,
                showlegend=True
            )
        )

    # Update the layout to include the markers and show legend
    fig_Brent.update_layout(showlegend=True)

    # Display the Plotly chart
    st.plotly_chart(fig_Brent, use_container_width=True, height=400)



    # Calculate Total_Position_Quantity and Weighted_Avg_Net_Premium for each portfolio
    filtered_df['Weighted_Avg_Net_Premium'] = (filtered_df['FO.NetPremium'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Protection'] = (filtered_df['FO.StrikePrice1'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Lower_Protection'] = (filtered_df['FO.StrikePrice2'] * filtered_df['FO.Position_Quantity']) / filtered_df.groupby('Portfolio')['FO.Position_Quantity'].transform('sum')
    filtered_df['Weighted_Avg_Protection_Band'] = filtered_df['Weighted_Avg_Protection'] - filtered_df['Weighted_Avg_Lower_Protection']
    filtered_df['Total_Cost'] = (filtered_df['FO.NetPremium'] * filtered_df['FO.Position_Quantity'])
    grouped_data = filtered_df.groupby('Portfolio').agg(
    Total_Position_Quantity=pd.NamedAgg(column='FO.Position_Quantity', aggfunc='sum'),
    Total_Cost=pd.NamedAgg(column='Total_Cost', aggfunc='sum'),
    Weighted_Avg_Net_Premium=pd.NamedAgg(column='Weighted_Avg_Net_Premium', aggfunc='sum'),
    Weighted_Avg_Protection=pd.NamedAgg(column='Weighted_Avg_Protection', aggfunc='sum'),
    Weighted_Avg_Lower_Protection=pd.NamedAgg(column='Weighted_Avg_Lower_Protection', aggfunc='sum'),
    Weighted_Avg_Protection_Band=pd.NamedAgg(column='Weighted_Avg_Protection_Band', aggfunc='sum'),
    Trade_Numbers=pd.NamedAgg(column='Portfolio', aggfunc='count')
    ).reset_index()

    # Apply accounting format to the numeric columns
    grouped_data['Total_Position_Quantity'] = grouped_data['Total_Position_Quantity'].apply('{:,.0f}'.format)
    grouped_data['Total_Cost'] = grouped_data['Total_Cost'].apply('USD{:,.2f}'.format)
    grouped_data['Weighted_Avg_Net_Premium'] = grouped_data['Weighted_Avg_Net_Premium'].apply('USD{:,.2f}'.format)
    grouped_data['Weighted_Avg_Protection'] = grouped_data['Weighted_Avg_Protection'].apply('USD{:,.2f}'.format)
    grouped_data['Weighted_Avg_Lower_Protection'] = grouped_data['Weighted_Avg_Lower_Protection'].apply('USD{:,.2f}'.format)
    grouped_data['Weighted_Avg_Protection_Band'] = grouped_data['Weighted_Avg_Protection_Band'].apply('USD{:,.2f}'.format)

    fig3 = go.Figure(data=[go.Table(
    header=dict(values=['Portfolio','Number of Trades', 'Total Volume Hedged','Total Cost', 'Weighted Average Net Premium','Weighted Average Protection','Weighted Average Lower Protection','Protection Band']),
    cells=dict(values=[grouped_data['Portfolio'], grouped_data['Trade_Numbers'],grouped_data['Total_Position_Quantity'],grouped_data['Total_Cost'], grouped_data['Weighted_Avg_Net_Premium']
                       , grouped_data['Weighted_Avg_Protection'], grouped_data['Weighted_Avg_Lower_Protection'], grouped_data['Weighted_Avg_Protection_Band']])
    )])

    # Add margin-bottom to reduce space after the table
    st.plotly_chart(fig3, use_container_width=True, height=200)

    st.divider()

    col1, col2 = st.columns((2))

    with col1:
        # Calculate Volume executed versus Counterparty
        st.subheader("Volume executed versus Counterparty")
        fig1 = px.histogram(filtered_df, x='FO.CounterpartyName', y='FO.Position_Quantity', color='FO.DealerID',title='Sum of Volume Executed',)
        fig1.update_xaxes(categoryorder='total descending')
        st.plotly_chart(fig1, use_container_width=True, height=200)
        

    with col1:
        df_refresh = pd.read_excel("PCHP Data.xlsx", "Sheet_Info")

        # Assuming 'Date_today' contains a single date value in the DataFrame
        date_today_value = df_refresh['Date_today'].iloc[0]

        # Convert to a datetime object and format it
        date_limit = datetime.strptime(str(date_today_value), "%Y-%m-%d %H:%M:%S.%f")

        # Format the datetime object to the desired format
        formatted_date_limit = date_limit.strftime("%d %b %Y")

        df_limits = pd.read_excel("PCHP Data.xlsx","Credit_Limit_data")
        st.subheader("Available Limits")
        df_limits = pd.read_excel("PCHP Data.xlsx","Credit_Limit_data")
        fig_limits = px.bar(df_limits, x='Counterparty', y=['Available Volume Limit', 'Volume Utilised'],
                title='Volume Limit and Volume Utilized by Counterparty as of '+ formatted_date_limit)
        #fig_limits .update_xaxes(categoryorder='total descending')
        st.plotly_chart(fig_limits, use_container_width=True, height=200)       

    with col2:
        st.subheader("Monthly Volume Executed")
        # Reshape the data to have 'Month' as a column and corresponding values
        df_melted = pd.melt(filtered_df, id_vars=['Portfolio'], value_vars=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                            var_name='Month', value_name='Value')

        # Define the correct order of months
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        # Convert 'Month' to a categorical data type with the correct order
        df_melted['Month'] = pd.Categorical(df_melted['Month'], categories=month_order, ordered=True)

        # Group by Portfolio, Month, and Value type (Quantity or Premium) and sum the values
        df_grouped = df_melted.groupby(['Portfolio', 'Month']).sum().reset_index()
    
        # Create a line chart for quantities
        fig_quantity = px.bar(df_grouped, x='Month', y='Value', color='Portfolio',
                            title='Quantity Comparison by Portfolio for Each Month',
                            labels={'Value': 'Quantity'}, barmode='group')



        # Add a horizontal line to indicate the targeted value
        default_targeted_value =  int(186480000 /12)  # Adjust this value according to your targeted value
        targeted_value = [default_targeted_value+85000,default_targeted_value-85000,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value,
                          default_targeted_value,default_targeted_value,default_targeted_value]

        # Create a trace for the target line
        target_trace = go.Scatter(x=df_grouped['Month'], y=[targeted_value],
                                mode='lines', line=dict(color='orange', dash='dash'),
                                name='FY2024 Mandated Volume')

        # Add the target trace to the figure
        fig_quantity.add_trace(target_trace)
        if len(selected_portfolios) == 1:
            # Calculate unexecuted volumes by subtracting executed volumes from the targeted value
            df_grouped['Unexecuted'] = targeted_value - df_grouped['Value']

            if selected_portfolios == ['FY2024 PCHP']:
                # Create a stacked bar chart with custom colors
                fig_stacked_bar = px.bar(df_grouped, x='Month', y=['Value', 'Unexecuted'],
                                        title='Executed vs. Unexecuted Volumes by Portfolio for Each Month',
                                        labels={'Value': 'Executed', 'Unexecuted': 'Unexecuted'},
                                        barmode='stack')
            else:
                # Create a stacked bar chart with custom colors
                fig_stacked_bar = px.bar(df_grouped, x='Month', y=['Value'], color='Portfolio',
                                        title='Executed vs. Unexecuted Volumes by Portfolio for Each Month',
                                        barmode='stack')

            # Set the color for "Unexecuted" bars to red
            fig_stacked_bar.update_traces(marker_color='red', selector=dict(name='Unexecuted'))

            st.plotly_chart(fig_stacked_bar, use_container_width=True, height=200)
        else:
            st.write("No data available for visualization.")

    with col2:
        st.subheader("Counterparty Monthly Volume Executed")
        # Reshape the data to have 'Month' as a column and corresponding values
        df_melted = pd.melt(filtered_df, id_vars=['FO.CounterpartyName'], value_vars=['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                            var_name='Month', value_name='Value')

        # Define the correct order of months
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        # Convert 'Month' to a categorical data type with the correct order
        df_melted['Month'] = pd.Categorical(df_melted['Month'], categories=month_order, ordered=True)

        # Group by Portfolio, Month, and Value type (Quantity or Premium) and sum the values
        df_grouped = df_melted.groupby(['FO.CounterpartyName', 'Month']).sum().reset_index()

        # Create a line chart for quantities
        fig_quantity = px.bar(df_grouped, x='Month', y='Value', color='FO.CounterpartyName',
                            title='Quantity Comparison by Counterparty for Each Month',
                            labels={'Value': 'Quantity'})

        st.plotly_chart(fig_quantity, use_container_width=True, height=200)


# Assuming selected_portfolio is a list
if len(selected_portfolios) > 0 and 'All' not in selected_portfolios:
    # Ensure only one element in selected_portfolio
    selected_portfolio = [selected_portfolios[0]]

def visualize_data(st, filtered_df, strike_price_column, strike_price_name):
    if not filtered_df.empty:
        # Remove rows with NaN values in the "FO.TransactionNumber" column
        filtered_df = filtered_df.dropna(subset=['FO.TransactionNumber'])

        # Remove rows with NaN values in the "Total Outstanding" column
        filtered_df = filtered_df.dropna(subset=['Total Outstanding'])

        # Check if "Total Outstanding" column is not empty
        if not filtered_df['Total Outstanding'].empty:
            # Remove rows with 0 values in the "Total Outstanding" column
            filtered_df = filtered_df[filtered_df['Total Outstanding'] != 0]

            # Extract relevant columns for visualization
            months = ['E.January', 'E.February', 'E.March', 'E.April', 'E.May', 'E.June', 'E.July', 'E.August', 'E.September', 'E.October', 'E.November', 'E.December']
            monthly_data = filtered_df[months]

            # Group by strike_price_column and sum the data
            grouped_data = filtered_df.groupby(strike_price_column)[months].sum()

            # Check if grouped_data is not empty
            if not grouped_data.empty:
                # Transpose the data for plotting
                transposed_data = grouped_data.transpose()

                # Plotting the data using Plotly
                fig = go.Figure()

                for col in transposed_data.columns:
                    fig.add_trace(go.Bar(x=transposed_data.index, y=transposed_data[col], name=col))

                fig.update_layout(
                    xaxis_title="Months",
                    yaxis_title="Total Barrels Executed",
                    xaxis_tickangle=-45,
                    barmode='stack',
                    legend=dict(title=strike_price_name, x=1, y=1)
                )

                # Display the Plotly chart
                st.plotly_chart(fig, use_container_width=True)

                # Display table
                st.subheader("Volume Breakdown")
                st.dataframe(grouped_data,height=150, use_container_width = True)

            else:
                st.write("No data available for visualization.")
        else:
            st.write("No data available for visualization. Total Outstanding column is empty.")
    else:
        st.write("No data available for visualization.")


def strike_data(st, filtered_df, strike_price_column, strike_price_name):
    if not filtered_df.empty:
        # Remove rows with NaN values in the "FO.TransactionNumber" column
        filtered_df = filtered_df.dropna(subset=['FO.TransactionNumber'])

        # Remove rows with NaN values in the "Total Outstanding" column
        filtered_df = filtered_df.dropna(subset=['Total Outstanding'])

        # Check if "Total Outstanding" column is not empty
        if not filtered_df['Total Outstanding'].empty:
            # Remove rows with 0 values in the "Total Outstanding" column
            filtered_df = filtered_df[filtered_df['Total Outstanding'] != 0]

            # Extract relevant columns for visualization
            months = ['E.January', 'E.February', 'E.March', 'E.April', 'E.May', 'E.June', 'E.July', 'E.August', 'E.September', 'E.October', 'E.November', 'E.December']
            monthly_data = filtered_df[months]

            # Group by strike_price_column and sum the data
            grouped_data = filtered_df.groupby(strike_price_column)[months].sum()

            # Check if grouped_data is not empty
            if not grouped_data.empty:
                # Transpose the data for plotting
                transposed_data = grouped_data.transpose()

                return grouped_data
                
with tab3:
    st.title("Mark to Market Data")
    col1, col2 = st.columns((2))
    
    # Check if "Total Outstanding" column is not empty
    if not filtered_df['Total Outstanding'].empty:
        with col1:
            st.subheader("Upper Strike Level")
            visualize_data(st, filtered_df, 'FO.StrikePrice1', 'FO.StrikePrice1')
        with col2:
            st.subheader("Lower Strike Level")
            visualize_data(st, filtered_df, 'FO.StrikePrice2', 'FO.StrikePrice2')
    else:
        st.write("No data available for visualization.")

    st.divider()
    st.title("BBG Option Price and Valuation")
    try:
        # Read the Excel file
        df_BBG = pd.read_excel("BBG_Output.xlsx", sheet_name=None)

        # Get all sheet names
        sheet_names = list(df_BBG.keys())

        # Create a dropdown to select sheet
        default_sheet = sheet_names[-1]  # Set the default value to the last sheet name
        selected_sheet = st.selectbox("Select a sheet", sheet_names, index=len(sheet_names)-1)

        # Show the selected sheet data
        st.write("Data Refreshed:", selected_sheet)

        # Rename the first column
        df_selected_sheet = df_BBG[selected_sheet].rename(columns={df_BBG[selected_sheet].columns[0]: 'Strike Price'})

        # Convert numerical values in the first column (except the last one) to integers with one decimal place
        for i in range(len(df_selected_sheet) - 1):
            value = df_selected_sheet.iloc[i, 0]
            if isinstance(value, (int, float)):
                df_selected_sheet.iloc[i, 0] = round(float(value), 1)

        # Convert the rounded numerical values to integers
        df_selected_sheet.iloc[:-1, 0] = df_selected_sheet.iloc[:-1, 0].astype(int)

        st.dataframe(df_selected_sheet, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error: {e}")

    st.divider()

    def process_dataframe(df1, df2):
        # Find common values in the first column of both dataframes
        common_values = df1.iloc[:, 0].isin(df2.iloc[:, 0])
        
        # Filter df1 and df2 based on common values in the first column
        df1_filtered = df1[df1.iloc[:, 0].isin(df2.iloc[:, 0])]
        df2_filtered = df2[df2.iloc[:, 0].isin(df1.iloc[:, 0])]
        
        df1_filtered = df1_filtered.reset_index(drop=True)
        df2_filtered = df2_filtered.reset_index(drop=True)
        # Reindex df2 to match the row and column indices of df1
        df2_reindexed = df2_filtered.reindex(index=df1_filtered.index, columns=df1_filtered.columns)
       
        # Multiply corresponding elements from df1 and df2
        result_df = df1_filtered * df2_reindexed

        # Assign the first column from df2 to the corresponding column in the result_df
        result_df[df1_filtered.columns[0]] = df1_filtered[df1_filtered.columns[0]]

        # Replace NaN values in result_df with 0
        result_df.fillna(0, inplace=True)

        return result_df

    
    # Process the first set of data
    df1 = df_selected_sheet
    df2 = strike_data(st, filtered_df, 'FO.StrikePrice1', 'FO.StrikePrice1')
    df2.reset_index(inplace=True)
    for column in df2.columns:
        df2[column] = df2[column].astype(int)
    df_Upper = process_dataframe(df1, df2)

    # Process the second set of data
    df1 = df_selected_sheet
    df3 = strike_data(st, filtered_df, 'FO.StrikePrice2', 'FO.StrikePrice2')
    df3.reset_index(inplace=True)
    for column in df3.columns:
        df3[column] = df3[column].astype(int)
    df_Lower = process_dataframe(df1, df3)

    # Display the results
    col3, col4 = st.columns((2))
    
    with col3:
        
        # Transpose the DataFrame to have months as columns and Strike Price as index
        df_Upper_transposed = df_Upper.set_index('Strike Price').transpose()

        # Create a Plotly bar chart
        fig = go.Figure()

        # Add bar trace for each Strike Price
        for i, strike_price in enumerate(df_Upper_transposed.columns):
            fig.add_trace(go.Bar(
                x=df_Upper_transposed.index,
                y=df_Upper_transposed[strike_price],
                name=f'Strike Price {strike_price}',
                marker_color=color_discrete_sequence[i % len(color_discrete_sequence)],
                text=df_Upper_transposed[strike_price],  # Use y-values as text
                textposition='outside',
                texttemplate='%{text:.2s}',
            ))

        # Update layout with axis labels and title
        fig.update_layout(xaxis_title='Tenure',
                        yaxis_title='Value, USD',
                        title='Valuation of Upper Put Options',legend=dict(x=0, y=1.0))

        # Show plot
        st.plotly_chart(fig)
        # Print DataFrame
        st.dataframe(df_Upper, height=150, use_container_width=True, hide_index=True)

        # Convert the chart to an image
        image = fig.to_image(format="png")

        # Save the image to a file
        image_path = r"Resources\Plots\upper_put_options.png"
        with open(image_path, "wb") as f:
            f.write(image)

        # Set up the file name
        filename = "plotly_chart.png"
        # Convert the image to bytes
        image_bytes_2 = io.BytesIO(image)
        # Trigger the download
        st.download_button(label="Download Image", data=image_bytes_2, file_name=filename, mime="image/png", key="download_button_1")

        
    with col4:
        # Transpose the DataFrame to have months as columns and Strike Price as index
        df_Lower_transposed = df_Lower.set_index('Strike Price').transpose()

        # Create a Plotly bar chart
        fig2 = go.Figure()


        # Define your discrete color sequence
        color_discrete_sequence = [
            "#00b1a9",  # Original color - R000 G177 B169
            "#763f98",  # Original color - R118 G063 B152
            "#20419a",  # Original color - R032 G065 B154
            "#fdb924",  # Original color - R253 G185 B036
            "#bfd730",  # Original color - R191 G215 B048
            "#007b73",  # Shade of R000 G177 B169
            "#3a1d4c",  # Shade of R118 G063 B152
            "#101e4a",  # Shade of R032 G065 B154
            "#cc8b1c",  # Shade of R253 G185 B036
            "#8e9c1b"   # Shade of R191 G215 B048
        ]


        # Add bar trace for each Strike Price
        for i, strike_price in enumerate(df_Lower_transposed.columns):
            fig2.add_trace(go.Bar(
                x=df_Lower_transposed.index,
                y=df_Lower_transposed[strike_price],
                name=f'Strike Price {strike_price}',
                marker_color=color_discrete_sequence[i % len(color_discrete_sequence)],
                text=df_Lower_transposed[strike_price],  # Use y-values as text
                textposition='outside',
                texttemplate='%{text:.2s}',
            ))

        # Update layout with axis labels and title
        fig2.update_layout(xaxis_title='Tenure',
                        yaxis_title='Value, USD',
                        title='Valuation of Lower Put Options',legend=dict(x=0, y=1.0))

        # Show plot
        st.plotly_chart(fig2)
    


        # Print DataFrame
        st.dataframe(df_Lower, height=150, use_container_width=True, hide_index=True)


        # Convert the chart to an image
        image = fig2.to_image(format="png")

        # Save the image to a file
        image_path = r"Resources\Plots\lower_put_options.png"
        with open(image_path, "wb") as f:
            f.write(image)
        # Set up the file name
        filename = "plotly_chart.png"
        # Convert the image to bytes
        image_bytes = io.BytesIO(image)

        

        # Trigger the download
        st.download_button(label="Download Image", data=image_bytes, file_name=filename, mime="image/png", key="download_button_2")

    st.divider()

     # Display PCHP Data
    st.title("Trade by trade Evaluation")

    # Create a formatted copy of the filtered DataFrame to preserve the original data
    formatted_df = filtered_df.copy()
    formatted_df_option = df_selected_sheet.copy()


        # Function to get list of months between two dates
    def get_months_between_dates(start_date, end_date):
        months = pd.date_range(start=start_date, end=end_date, freq='MS').strftime('%B').tolist()
        return months

    # Format date columns for better readability
    date_columns = ['FO.TradeDate', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate']
    for column in date_columns:
        formatted_df[column] = formatted_df[column].dt.strftime('%d %b %Y')

    # Create a new column 'MonthsBetween' containing list of months between start and end fix dates
    formatted_df['Tenure'] = formatted_df.apply(lambda row: get_months_between_dates(row['FO.StartFixDate'], row['FO.EndFixDate']), axis=1)

    # Assuming 'formatted_df' is your DataFrame
    formatted_df.rename(columns={'Row Labels':'Trade Number'}, inplace=True)

    # Specify columns to display in the table
    columns_to_display = ['Trade Number','Portfolio','FO.TradeDate','FO.DealerID', 'FO.CounterpartyName','FO.StructureType_label','FO.NetPremium', 'FO.Position_Quantity',
                        'FO.StrikePrice1', 'FO.StrikePrice2', 'FO.StartFixDate', 'FO.EndFixDate', 'FO.Settlement_DeliveryDate',
                        'E.January','E.February','E.March','E.April','E.May','E.June','E.July',
                        'E.August','E.September','E.November','E.December']

    # Reset index to start from 1
    formatted_df = formatted_df.reset_index(drop=True)

    # Start index from 1
    formatted_df.index = formatted_df.index + 1
    formatted_df_option_E = formatted_df.copy()

    # Find common months between both DataFrames
    common_months = [col for col in formatted_df_option.columns if col.startswith('E.')]

    # Iterate over each row in formatted_df
    for index, row in formatted_df.iterrows():
        # Get the Strike Price from formatted_df
        strike_price_1 = row['FO.StrikePrice1']
        strike_price_2 = row['FO.StrikePrice2']

        # Find corresponding row in formatted_df_option
        option_row_1 = formatted_df_option[formatted_df_option['Strike Price'] == strike_price_1]
        option_row_2 = formatted_df_option[formatted_df_option['Strike Price'] == strike_price_2]

        # Check if option_row is not empty
        if not option_row_1.empty and not option_row_2.empty:
            # Multiply the values in common months and update the row in formatted_df
            for month in common_months:
                formatted_df.at[index, month] = formatted_df.at[index, month] * option_row_1[month].iloc[0] - formatted_df.at[index, month] * option_row_2[month].iloc[0] 


    # Now the values in formatted_df are updated according to the conditions specified

    # List of columns related to the months
    month_columns = ['E.January', 'E.February', 'E.March', 'E.April', 'E.May', 'E.June', 'E.July', 'E.August', 'E.September', 'E.October','E.November', 'E.December']
    month_columns_value = ['January,USD', 'February,USD', 'March,USD', 'April,USD', 'May,USD', 'June,USD', 'July,USD', 'August,USD', 'September,USD', 'October,USD', 'November,USD', 'December,USD']


    

    # Assuming 'formatted_df' is your DataFrame
    formatted_df.rename(columns={
        'E.January': 'January,USD',
        'E.February': 'February,USD',
        'E.March': 'March,USD',
        'E.April': 'April,USD',
        'E.May': 'May,USD',
        'E.June': 'June,USD',
        'E.July': 'July,USD',
        'E.August': 'August,USD',
        'E.September': 'September,USD',
        'E.October': 'October,USD',
        'E.November': 'November,USD',
        'E.December': 'December,USD'
    }, inplace=True)


    
    # Create new columns with default value 0 in formatted_df
    for col in month_columns:
        formatted_df[col] = 0

    # Assign values from formatted_df_option_E to formatted_df
    formatted_df[month_columns] = formatted_df_option_E[month_columns]  
    columns_to_display.extend(month_columns_value)

    formatted_df.rename(columns={
        'E.January': 'January,bbls',
        'E.February': 'February,bbls',
        'E.March': 'March,bbls',
        'E.April': 'April,bbls',
        'E.May': 'May,bbls',
        'E.June': 'June,bbls',
        'E.July': 'July,bbls',
        'E.August': 'August,bbls',
        'E.September': 'September,bbls',
        'E.October': 'October,bbls',
        'E.November': 'November,bbls',
        'E.December': 'December,bbls'
    }, inplace=True)

    columns_to_display = [f"{column.split('.')[1]},bbls" if column.startswith('E.') else column for column in columns_to_display]


    # Assuming you have a DataFrame named 'data' containing your dataset
    formatted_df['Value at inception'] = formatted_df['FO.NetPremium'] * formatted_df['FO.Position_Quantity']
    columns_to_display.append('Value at inception')

    # Add a new column 'Total' containing the sum of values in the month columns
    formatted_df['Current Value'] = formatted_df[month_columns_value].sum(axis=1)
    columns_to_display.append('Current Value')

    # Now the values in formatted_df_option are updated according to the conditions specified
    with st.container():
        # Show the formatted DataFrame using st.dataframe
        st.dataframe(formatted_df[columns_to_display], height=500 ,use_container_width=True)

    # buffer to use for excel writer
    buffer = io.BytesIO()

    # Download Button
    @st.cache_data
    def convert_to_excel(formatted_df, df_selected_sheet):
        # Create Excel writer object
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write formatted_df to the first sheet
            formatted_df.to_excel(writer, sheet_name='Portfolio Sum', index=False)

            # Write df_selected_sheet to the second sheet
            df_selected_sheet.to_excel(writer, sheet_name='Option Data', index=False)

            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet1 = writer.sheets['Portfolio Sum']
            worksheet2 = writer.sheets['Option Data']

            # Define cell formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'border': 1,
                'font_color': 'white',  # Set font color to white
                'bg_color': '#38B09D'  # Set background color to 38B09D
            })
            data_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1})

            # Apply formatting to first sheet (Portfolio Sum)
            for col_num, value in enumerate(formatted_df.columns.values):
                worksheet1.write(0, col_num, value, header_format)
                worksheet1.set_column(col_num, col_num, 15, data_format)  # Set column width to 100 pixels

            # Apply formatting to second sheet (Option Data)
            for col_num, value in enumerate(df_selected_sheet.columns.values):
                worksheet2.write(0, col_num, value, header_format)
                worksheet2.set_column(col_num, col_num, 15, data_format)  # Set column width to 100 pixels

             # Set row height in points (1 point ≈ 0.75 pixels)
            row_height_in_points = 50
            worksheet1.set_default_row(row_height_in_points)  # Set default row height for the first sheet
            worksheet2.set_default_row(row_height_in_points)  # Set default row height for the second sheet

            # Freeze the header row
            worksheet1.freeze_panes(1, 0)  # Freeze the first row in the first sheet

            # Close the Pandas Excel writer
            writer.close()

        return buffer.getvalue()


    excel_data = convert_to_excel(formatted_df[columns_to_display], df_selected_sheet)

    # Get today's date and format it
    today_date = datetime.now().strftime('%d-%m-%Y')

    # Download button to download Excel file
    download_button = st.download_button(
        label="Download data as Excel",
        data=excel_data,
        file_name=f"{today_date}_MTM.xlsx",  # Set file name with today's date
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

with tab4:
    
    
    def create_download_link(val, filename):
        b64 = base64.b64encode(val)  # val looks like b'...'
        return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}.pdf">Download file</a>'

    def create_letterhead(pdf, WIDTH):
        pdf.image(r"Resources/Blue Modern Business Letterhead.jpg", 0, 0, WIDTH)

    def create_title(title, pdf):
        # Add main title
        pdf.set_font('Helvetica', 'b', 20)  
        pdf.ln(40)
        pdf.write(5, title)
        pdf.ln(10)
        # Add date of report
        pdf.set_font('Helvetica', '', 14)
        pdf.set_text_color(r=128,g=128,b=128)
        today = time.strftime("%d/%m/%Y")
        pdf.write(4, f'{today}')
        # Add line break
        pdf.ln(10)

    def write_to_pdf(pdf, words):
        # Set text colour, font size, and font type
        pdf.set_text_color(r=0,g=0,b=0)
        pdf.set_font('Helvetica', '', 12)
        pdf.write(5, words)

    class PDF(FPDF):
        def footer(self):
            self.set_y(-15)
            self.set_font('Helvetica', 'I', 8)
            self.set_text_color(128)
            self.cell(0, 10, 'Page ' + str(self.page_no()), 0, 0, 'C')

    # Global Variables
    TITLE = "Monthly Business Report"
    WIDTH = 210
    HEIGHT = 297

    # Create PDF
    pdf = PDF() # A4 (210 by 297 mm)

    # Add Page
    pdf.add_page()

    # Add lettterhead and title
    create_letterhead(pdf, WIDTH)
    create_title(TITLE, pdf)

    # Add some content to the PDF
    content = [
        "1. The table below illustrates the annual sales of Heicoders Academy:",
        "2. The visualisations below show the trend of total sales for Heicoders Academy and the breakdown of revenue for year 2016:"
    ]

    for item in content:
        write_to_pdf(pdf, item)
        pdf.ln(15)

    
    pdf.ln(10)

    # Add Page
    pdf.add_page()

    # Add lettterhead
    create_letterhead(pdf, WIDTH)
    create_title(TITLE, pdf)

    # Add dynamically generated image to the PDF
    # Assuming image_bytes contains the bytes of the image generated by Plotly
    pdf.image(r"Resources\Plots\upper_put_options.png", x=5, y=pdf.get_y(), w=100)
    pdf.ln(15)
    pdf.image(r"Resources\Plots\lower_put_options.png", x=100, y=pdf.get_y(), w=100) 
    

    # Generate the PDF and provide download link
    pdf_output = pdf.output(dest="S").encode("latin-1")
    html = create_download_link(pdf_output, "report")
    st.markdown(html, unsafe_allow_html=True)

