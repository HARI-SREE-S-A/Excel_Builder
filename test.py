import dash
from dash import dcc, html, Input, Output
from dash.dash_table import DataTable
import pandas as pd
import plotly.express as px
import openpyxl

# Read Excel file into a pandas DataFrame
file_path = 'Call Entries updated.xlsx'

# Define specific sheet names
sheet_names = ["Kavitha", "Meenu", "Ajanya", "AJITH"]

# Create a dictionary to store category data for each sheet
category_data = {}

# Consolidate data from all sheets into a single DataFrame
consolidated_data = pd.DataFrame()
for sheet_name in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    # Standardize column names to ensure consistency
    if 'Category' in df.columns:
        df.rename(columns={'Category': 'Category'}, inplace=True)
    elif 'Category ' in df.columns:
        df.rename(columns={'Category ': 'Category'}, inplace=True)
    elif 'Category:' in df.columns:
        df.rename(columns={'Category:': 'Category'}, inplace=True)
    else:
        raise ValueError(f"Category column not found in sheet '{sheet_name}'")

    category_data[sheet_name] = df
    consolidated_data = pd.concat([consolidated_data, df], ignore_index=True)

# Initialize the Dash app
app = dash.Dash(__name__)

# Define the layout of the Dash app
app.layout = html.Div([
    html.H1("Category Distribution"),

    # Display pie charts for each sheet
    html.Div([
        dcc.Graph(id=f'pie-chart-{sheet_name}', figure=px.pie(
            category_data[sheet_name].groupby('Category').size().reset_index(name='Count'),
            values='Count', names='Category', title=f'Category Distribution - {sheet_name}'
        ))
        for sheet_name in category_data.keys()
    ]),

    # Display pie chart for consolidated data
    html.Div([
        dcc.Graph(id='consolidated-pie-chart')
    ]),

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


# Define callback to update the table based on pie chart selection
@app.callback(
    Output('table', 'data'),
    [Input(f'pie-chart-{sheet_name}', 'clickData') for sheet_name in category_data.keys()] +
    [Input('consolidated-pie-chart', 'clickData')]
)
def update_table(*clickData):
    selected_category = None
    for data in clickData:
        if data:
            selected_category = data['points'][0]['label']
            break

    if selected_category is None:
        return []

    if 'consolidated-pie-chart' in dash.callback_context.triggered_id:
        # Filter data based on selected category for consolidated data
        filtered_data = consolidated_data[consolidated_data['Category'] == selected_category]
    else:
        # Find the sheet name associated with the clicked pie chart
        clicked_sheet = None
        for sheet_name in category_data.keys():
            if f'pie-chart-{sheet_name}' in dash.callback_context.triggered_id:
                clicked_sheet = sheet_name
                break

        if clicked_sheet is None:
            return []

        # Filter data based on selected category and clicked sheet
        filtered_data = category_data[clicked_sheet][category_data[clicked_sheet]['Category'] == selected_category]

    return filtered_data.to_dict('records')


# Callback to update consolidated pie chart
@app.callback(
    Output('consolidated-pie-chart', 'figure'),
    [Input(f'pie-chart-{sheet_name}', 'clickData') for sheet_name in category_data.keys()]
)
def update_consolidated_pie_chart(*clickData):
    category_counts = consolidated_data.groupby('Category').size().reset_index(name='Count')
    return px.pie(category_counts, values='Count', names='Category', title='Consolidated Category Distribution')


# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)
