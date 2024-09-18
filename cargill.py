import streamlit as st
import pandas as pd
import plotly.express as px

# Custom CSS for thinner horizontal line
st.markdown("""
    <style>
    .horizontal-line {
        border-top: 1px solid #000000;  /* Thinner line */
        margin: 20px 0;
    }
    </style>
""", unsafe_allow_html=True)

# Title and subtitle
st.markdown("""
    <div style="font-size: 40px; font-weight: bold; color: black; text-align: center; margin-bottom: 10px;">
        Main Dashboard
    </div>
    <div style="font-size: 16px; text-align: center; color: grey; margin-bottom: 30px;">
        Cost Analysis Proyek Cargill
    </div>
""", unsafe_allow_html=True)

# Horizontal line to separate sections
st.markdown('<div class="horizontal-line"></div>', unsafe_allow_html=True)

# Sidebar for file upload and filters
st.sidebar.header("Upload and Filters")

# Upload Excel file in the sidebar
uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # Ensure the date column is in datetime format
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')

    # Sidebar filters for Date Range, WBS, and Cost Code
    st.sidebar.header('Filter Options')
    
    # Date Range Filter
    start_date = st.sidebar.date_input('Start Date', df['DATE'].min())
    end_date = st.sidebar.date_input('End Date', df['DATE'].max())
    
    # WBS Filter
    selected_wbs = st.sidebar.multiselect('Select WBS', options=df['WBS'].unique(), default=df['WBS'].unique())
    
    # Cost Code Filter
    selected_cost_code = st.sidebar.multiselect('Select Cost Code', options=df['COST CODE'].unique(), default=df['COST CODE'].unique())
    
    # Filter dataframe based on selections
    filtered_df = df[(df['DATE'] >= pd.to_datetime(start_date)) & (df['DATE'] <= pd.to_datetime(end_date))]
    filtered_df = filtered_df[filtered_df['WBS'].isin(selected_wbs) & filtered_df['COST CODE'].isin(selected_cost_code)]
    
    # Sort the DataFrame by DATE to ensure chronological order
    filtered_df = filtered_df.sort_values(by='DATE')

    # Add cumulative sum for Amount column
    filtered_df['Cumulative Amount'] = filtered_df.groupby('WBS')['AMOUNT'].cumsum()

    # Divs for chart and table, initially showing the chart
    st.markdown('<div id="chart-div">', unsafe_allow_html=True)

    # Cumulative line chart using Plotly
    fig = px.line(filtered_df, x='DATE', y='Cumulative Amount', color='WBS',
                  labels={'DATE': 'Date', 'Cumulative Amount': 'Cumulative Amount'}, 
                  title='Cumulative Amount Over Time by WBS',
                  markers=True)

    # Thin line with smooth curve
    fig.update_traces(mode="lines+markers", line_shape='spline', line_smoothing=0.6, line=dict(width=2))

    # Format currency for annotations
    def format_currency(value):
        return f"Rp {value:,.0f}".replace(',', '.')

    # Highlight last value with Rp format
    last_values = filtered_df.groupby('WBS').tail(1)
    for i, row in last_values.iterrows():
        fig.add_annotation(
            x=row['DATE'],
            y=row['Cumulative Amount'],
            text=format_currency(row['Cumulative Amount']),
            showarrow=True,
            arrowhead=2,
            ax=-20,
            ay=-40,
            font=dict(size=12, color="#000000")
        )

    # Update x-axis to display dates from March 2023 to October 2023
    fig.update_xaxes(
        range=[pd.Timestamp('2023-03-01'), pd.Timestamp('2023-10-31')],
        tickformat="%b %Y",  # Format x-axis ticks as "Month Year"
        title_text='Date'
    )

    # Update y-axis to display values with Rp currency format
    fig.update_yaxes(tickprefix="Rp ")

    # Display the chart
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('</div><div id="table-div">', unsafe_allow_html=True)

    # Filter data from March 2023 to October 2023 for table view
    table_filtered_df = filtered_df[(filtered_df['DATE'] >= '2023-03-01') & (filtered_df['DATE'] <= '2023-10-31')]

    # Define function to calculate monthly amounts based on actual values
    def calculate_period_summary(df, period_start, period_end):
        return df[(df['DATE'] >= period_start) & (df['DATE'] <= period_end)].groupby(['WBS', 'COST CODE'])['AMOUNT'].sum()

    # Calculate Previous Month (before start_date), This Month (within start_date to end_date), Next Month (after end_date)
    previous_month = calculate_period_summary(df, df['DATE'].min(), pd.to_datetime(start_date) - pd.DateOffset(days=1))  # Before start date
    this_month = calculate_period_summary(df, pd.to_datetime(start_date), pd.to_datetime(end_date))  # Between start_date and end_date
    next_month = calculate_period_summary(df, pd.to_datetime(end_date) + pd.DateOffset(days=1), df['DATE'].max())  # After end date

    # Create summary DataFrame with both 'WBS' and 'COST CODE'
    summary_df = pd.DataFrame({
        'WBS': this_month.index.get_level_values(0),  # WBS codes
        'COST CODE': this_month.index.get_level_values(1),  # Cost Codes
        'Previous_Month': previous_month.reindex(this_month.index, fill_value=0).values,
        'This_Month': this_month.values,
        'Next_Month': next_month.reindex(this_month.index, fill_value=0).values
    })

    # Calculate total row for each period
    total_row = summary_df[['Previous_Month', 'This_Month', 'Next_Month']].sum()
    total_row_df = pd.DataFrame([['Total', '', total_row['Previous_Month'], total_row['This_Month'], total_row['Next_Month']]], 
                                columns=['WBS', 'COST CODE', 'Previous_Month', 'This_Month', 'Next_Month'])
    
    # Append total row
    summary_df = pd.concat([summary_df, total_row_df], ignore_index=True)

    # Format currency for table
    summary_df['Previous_Month'] = summary_df['Previous_Month'].apply(lambda x: f"Rp {x:,.0f}".replace(',', '.'))
    summary_df['This_Month'] = summary_df['This_Month'].apply(lambda x: f"Rp {x:,.0f}".replace(',', '.'))
    summary_df['Next_Month'] = summary_df['Next_Month'].apply(lambda x: f"Rp {x:,.0f}".replace(',', '.'))

    # Display the summarized data table with adjusted column widths
    st.dataframe(summary_df.style.set_properties(subset=['Previous_Month', 'This_Month', 'Next_Month'], **{'width': '150px'}), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.write("Please upload an Excel file to begin.")
