from collections import namedtuple
import altair as alt
import math
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import datetime as dt
from datetime import datetime





# Setting Up
st.set_page_config(page_title = "DashBoard",page_icon = '4953098.png',layout ="wide")

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
tab1, tab2 = st.tabs(["Market", "Execution data"])


# Get the size of the primary monitor
default_height = 540
default_width = 1056
# DatMarket chart sizee
st.sidebar.header("Market Chart Size")
height = st.sidebar.slider("Height", 200, 1500,default_height , 50)
width = st.sidebar.slider("Width", 200, 1500, default_width, 50)

with tab1:
    # Ticker Tape
    st.subheader("Live Price and News")
    st.components.v1.html(
        """
        <style>
            .tradingview-widget-container {
                background-color: transparent !important;
            }
        </style>
        <!-- TradingView Widget BEGIN -->
        <div class="tradingview-widget-container">
            <div class="tradingview-widget-container__widget"></div>
            <div class="tradingview-widget-copyright">
                <a href="https://www.tradingview.com" rel="noopener" target="_blank">
                    <span class="blue-text"></span>
                </a>
            </div>
            <script type="text/javascript" src="https://s3.tradingview.com/external-embedding/embed-widget-ticker-tape.js" async>
            {
                "symbols": [
                    {
                        "proName": "VELOCITY:BRENT",
                        "title": "Spot Brent"
                    },
                    {
                        "proName": "CAPITALCOM:DXY",
                        "title": "Dollar Index"
                    },
                    {
                        "proName": "FOREXCOM:SPXUSD",
                        "title": "S&P 500"
                    },
                    {
                        "proName": "FOREXCOM:NSXUSD",
                        "title": "Nasdaq 100"
                    }
                ],
                "colorTheme": "dark",
                "isTransparent": false,
                "displayMode": "adaptive",
                "locale": "en"
            }
            </script>
        </div>
        <!-- TradingView Widget END -->
        """,
        width=None,
        height=None,
        scrolling=False,
    )

    col1,col2= st.columns((2))

    with col1:
        st.subheader("Brent Spot Price")
        st.components.v1.html(f"""<!-- TradingView Widget BEGIN -->
        <div class="tradingview-widget-container">
        <div id="tradingview_36ce6"></div>
        <div class="tradingview-widget-copyright"><a href="https://www.tradingview.com/" rel="noopener nofollow" target="_blank"><span class="blue-text"></div>
        <script type="text/javascript" src="https://s3.tradingview.com/tv.js"></script>
        <script type="text/javascript">
        new TradingView.widget(
        {{
        "width": {width},
        "height": {height},
        "symbol": "VELOCITY:BRENT",
        "interval": "15",
        "timezone": "Asia/Hong_Kong",
        "theme": "dark",
        "style": "1",
        "locale": "en",
        "enable_publishing": false,
        "hide_legend": true,
        "withdateranges": true,
        "container_id": "tradingview_36ce6"
        }}
        );
        </script>
        </div>
        <!-- TradingView Widget END -->""",
        width=width, height=height, scrolling=False)
    

        st.subheader("Economic Data")
        st.components.v1.html(f"""<!-- TradingView Widget BEGIN -->
        <div class="tradingview-widget-container">
        <div class="tradingview-widget-container__widget"></div>
        <div class="tradingview-widget-copyright"><a href="https://www.tradingview.com/" rel="noopener nofollow" target="_blank"><span class="blue-text"></div>
        <script type="text/javascript" src="https://s3.tradingview.com/external-embedding/embed-widget-events.js" async>
        {{
        "width": {width},
        "height": {height},
        "colorTheme": "dark",
        "isTransparent": false,
        "locale": "en",
        "importanceFilter": "-1,0,1",
        "currencyFilter": "USD,CNY,EUR,MYR,GBP"
        }}
        </script>
        </div>
        <!-- TradingView Widget END -->""",
                    
                    width=width, height=height, scrolling=False)




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
