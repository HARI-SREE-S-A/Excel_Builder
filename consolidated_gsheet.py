import dash
from dash import dcc, html, Input, Output
from dash.dash_table import DataTable
import pandas as pd
import plotly.express as px
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define Google Sheets credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = {
  "type": "serviaccount",
  "project_id": "excel-report-918",
  "private_key_id": "a85d86967840702f228fad9d7919dcb2",
  "private_key": "-----BEGIN PRIVATE KEY-----\n\n-----END PRIVATE KEY-----\n",
  "client_email": "excel-api@excel-report-419918.iam.gserviceaccount.com",
  "client_id": "104174658286889668287",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googlis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.gois.com/robot/v1/tadata/x509/excel-api%4xcel-rport1918.iamseviceaccount.com",
  "universe_domain": "googleapis.com"
}

# Authenticate with Google Sheets API
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials, scope)
gc = gspread.authorize(credentials)

# Open the Google Sheet
sheet_url = 'https://docs.google.com/spreadsheets/d/1LZ7-HhZddRrOLraTUWLTQH_dyGiQC7Bya2ssnZucChs/edit?usp=drive_link'
sh = gc.open_by_url(sheet_url)

# Read specific sheets into a pandas DataFrame
sheet_names = ["Kavitha", "Meenu", "Ajanya", "AJITH"]
category_data = {}
consolidated_data = pd.DataFrame()

for sheet_name in sheet_names:
    worksheet = sh.worksheet(sheet_name)
    records = worksheet.get_all_records(expected_headers=["Date", "Category"])
    df = pd.DataFrame(records)
    if 'Category' in df.columns:
        df.rename(columns={'Category': 'Category'}, inplace=True)
    elif 'Category ' in df.columns:
        df.rename(columns={'Category ': 'Category'}, inplace=True)
    elif 'Category:' in df.columns:
        df.rename(columns={'Category:': 'Category'}, inplace=True)
    else:
        raise ValueError(f"Category column not found in sheet '{sheet_name}'")
    # Parse the 'Date' column as datetime with the 'mixed' format option
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')
    category_data[sheet_name] = df
    consolidated_data = pd.concat([consolidated_data, df], ignore_index=True)

# Initialize the Dash app
app = dash.Dash(__name__)

# Define the layout of the Dash app
app.layout = html.Div([
    html.H1("Category Distribution"),

    # Date Picker for the dynamic pie chart
    dcc.DatePickerSingle(
        id='date-picker',
        min_date_allowed=consolidated_data['Date'].min(),
        max_date_allowed=consolidated_data['Date'].max(),
        initial_visible_month=consolidated_data['Date'].max(),
        date=consolidated_data['Date'].max()  # Default to the latest date
    ),

    # First pie chart showing consolidated data count for category column in all four sheets
    dcc.Graph(id='consolidated-pie-chart'),

    # Second pie chart which will be dynamic based on selected date
    dcc.Graph(id='dynamic-pie-chart'),

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
            row_selectable='single'  # Allow selecting a single row
        )
    ])
])

# Callback to update the second pie chart and table based on selected date
@app.callback(
    [Output('dynamic-pie-chart', 'figure'),
     Output('table', 'data')],
    [Input('date-picker', 'date'),
     Input('dynamic-pie-chart', 'clickData')]  # Add input for clickData
)
def update_visuals(selected_date, clickData):
    filtered_consolidated_data = consolidated_data[consolidated_data['Date'] == selected_date]

    pie_charts = []
    table_data = []

    for sheet_name, df in category_data.items():
        filtered_sheet_data = df[df['Date'] == selected_date]
        pie_chart = dcc.Graph(
            id=f'pie-chart-{sheet_name}',
            figure=px.pie(filtered_sheet_data.groupby('Category').size().reset_index(name='Count'),
                          values='Count', names='Category', title=f'Category Distribution - {sheet_name}')
        )
        pie_charts.append(pie_chart)

        table_data.extend(filtered_sheet_data.to_dict('records'))

    # Dynamic pie chart based on selected date
    dynamic_pie_chart_figure = px.pie(filtered_consolidated_data.groupby('Category').size().reset_index(name='Count'),
                                      values='Count', names='Category', title='Dynamic Category Distribution')

    # Update table data based on clicked category
    if clickData:
        clicked_category = clickData['points'][0]['label']
        table_data = [row for row in table_data if row['Category'] == clicked_category]

    return dynamic_pie_chart_figure, table_data

# Callback to update the first pie chart for consolidated data
@app.callback(
    Output('consolidated-pie-chart', 'figure'),
    [Input('date-picker', 'date')]
)
def update_consolidated_pie_chart(selected_date):
    # Static pie chart for consolidated data (total count)
    consolidated_pie_chart_figure = px.pie(consolidated_data.groupby('Category').size().reset_index(name='Count'),
                                           values='Count', names='Category', title='Consolidated Category Distribution')
    return consolidated_pie_chart_figure


