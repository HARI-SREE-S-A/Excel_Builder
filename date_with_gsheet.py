import dash
from dash import dcc, html, Input, Output
from dash.dash_table import DataTable
import pandas as pd
import plotly.express as px
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Set up the Google Sheet connection
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('path/to/your/credentials.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheet
sheet = client.open("Call Entries updated").sheet1

# Read the data into a pandas DataFrame
data = sheet.get_all_values()
column_names = data.pop(0)
df = pd.DataFrame(data, columns=column_names)

# Define specific sheet names
sheet_names = ["Kavitha", "Meenu", "Ajanya", "AJITH"]
# Create a dictionary to store category data for each sheet
category_data = {}
# Consolidate data from all sheets into a single DataFrame
consolidated_data = pd.DataFrame()
for sheet_name in sheet_names:
    sheet_data = df[df['Sheet'] == sheet_name]
    category_data[sheet_name] = sheet_data
    consolidated_data = pd.concat([consolidated_data, sheet_data], ignore_index=True)

# Initialize the Dash app
app = dash.Dash(__name__)

# Define the layout of the Dash app
app.layout = html.Div([
    html.H1("Category Distribution"),
    # Date Picker
    dcc.DatePickerSingle(
        id='date-picker',
        min_date_allowed=consolidated_data['Date'].min(),
        max_date_allowed=consolidated_data['Date'].max(),
        initial_visible_month=consolidated_data['Date'].max(),
        date=consolidated_data['Date'].max()  # Default to the latest date
    ),
    # Display pie charts for each sheet
    html.Div(id='pie-charts', className='pie-chart-container'),
    # Display the interactive table
    html.Div([
        html.H2("Category Wise Data"),
        DataTable(
            id='table',
            columns=[{"name": i, "id": i} for i in category_data[next(iter(category_data))].columns],
            data=[],
            export_format='csv',  # Enable exporting to CSV
            sort_action='native',  # Enable native sorting
            filter_action='native',  # Enable native filtering
            page_action='native',  # Enable pagination
            page_size=10,  # Set number of rows per page
        )
    ])
])

# Callback to update pie charts and table based on selected date and pie chart click
@app.callback(
    [Output('table', 'data'), Output('table', 'columns')],
    [Input(f'pie-chart-{sheet_name}', 'clickData') for sheet_name in sheet_names],
    [Input('date-picker', 'date')])
def update_table(kavitha_click, meenu_click, ajanya_click, ajith_click, selected_date):
    pie_chart_click_data = {
        "Kavitha": kavitha_click,
        "Meenu": meenu_click,
        "Ajanya": ajanya_click,
        "AJITH": ajith_click
    }
    table_data = []
    table_columns = []

    for sheet_name, df in category_data.items():
        if pie_chart_click_data[sheet_name]:
            # Get the category from the clicked pie chart
            category = pie_chart_click_data[sheet_name]['points'][0]['label']

            # Filter the data based on the selected date and category
            filtered_data = df[df['Date'] == selected_date]
            filtered_data = filtered_data[filtered_data['Category'] == category]

            # Sort the data by date
            filtered_data = filtered_data.sort_values('Date')

            # Update the table data and columns
            table_data = filtered_data.to_dict('records')
            table_columns = [{"name": i, "id": i} for i in filtered_data.columns]
            break
    else:
        table_data = []
        table_columns = [{"name": i, "id": i} for i in category_data[next(iter(category_data))].columns]

    return table_data, table_columns

# Callback to create the pie charts
@app.callback(
    Output('pie-charts', 'children'),
    [Input('date-picker', 'date')])
def update_pie_charts(selected_date):
    pie_charts = []
    for sheet_name, df in category_data.items():
        filtered_sheet_data = df[df['Date'] == selected_date]
        pie_chart = dcc.Graph(
            id=f'pie-chart-{sheet_name}',
            figure=px.pie(filtered_sheet_data.groupby('Category').size().reset_index(name='Count'),
                         values='Count', names='Category', title=f'Category Distribution - {sheet_name}'),
            config={'displayModeBar': False}
        )
        pie_charts.append(pie_chart)
    return pie_charts

# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)
