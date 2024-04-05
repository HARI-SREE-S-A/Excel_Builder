import dash
from dash import dcc, html, Input, Output
from dash.dash_table import DataTable
import pandas as pd
import plotly.express as px



# Read Excel file into a pandas DataFrame
file_path = 'data.xlsx'
xls = pd.ExcelFile(file_path, engine='openpyxl')

# Create a dictionary to store category data for each sheet
category_data = {}

# Iterate through each sheet in the Excel file
for sheet_name in xls.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    # Check if 'Category' column exists
    if 'Category' in df.columns:
        # Store category data in the dictionary
        category_data[sheet_name] = df

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
    [Input(f'pie-chart-{sheet_name}', 'clickData') for sheet_name in category_data.keys()]
)
def update_table(*clickData):
    selected_category = None
    for data in clickData:
        if data:
            selected_category = data['points'][0]['label']
            break

    if selected_category is None:
        return []

    # Filter data based on selected category
    filtered_data = pd.concat([df[df['Category'] == selected_category] for df in category_data.values()])

    return filtered_data.to_dict('records')

# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)
