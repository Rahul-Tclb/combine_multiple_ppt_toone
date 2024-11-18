import pandas as pd
import xlsxwriter

def create_dynamic_bar_chart(filename, output_excel):
    """
    Creates a clustered column chart in Excel using xlsxwriter.
    Automatically updates grouped data and chart when original data changes.

    Args:
        filename: The name of the input Excel file.
        output_excel: The name of the output Excel file with the chart.
    """
    # Load the data
    df = pd.read_excel(filename)
    client_name = df['ClientName'].unique()[0]

    # Group the data by 'ReportMonth' and 'Status' and calculate the sum of 'DistinctCases'
    grouped_df = df.groupby(['ReportMonth', 'Status']).sum().reset_index()

    # Create a pivot table for easier plotting
    pivot_df = grouped_df.pivot_table(index='ReportMonth', columns='Status', values='DistinctCases', fill_value= 0)
    pivot_df.reset_index(inplace=True)  # Ensure 'ReportMonth' is a column

    # Create an Excel file
    workbook = xlsxwriter.Workbook(output_excel)
    
    # Write original data to a sheet
    sheet_original = 'Original Data'
    worksheet_original = workbook.add_worksheet(sheet_original)
    worksheet_original.write_row(0, 0, df.columns)  # Write headers
    for row_num, row in enumerate(df.values, start=1):
        worksheet_original.write_row(row_num, 0, row)

    # Write grouped data to a new sheet
    sheet_grouped = 'Grouped Data'
    worksheet_grouped = workbook.add_worksheet(sheet_grouped)
    worksheet_grouped.write_row(0, 0, ['ReportMonth'] + list(pivot_df.columns[1:]))  # Write headers
    for row_num, row in enumerate(pivot_df.values, start=1):
        worksheet_grouped.write_row(row_num, 0, row)

    # Create a clustered column chart (without 'subtype': 'stacked')
    chart = workbook.add_chart({'type': 'column'})

    # Add data series dynamically linked to the grouped data
    for col_num in range(1, pivot_df.shape[1]):
        chart.add_series({
            'name':       [sheet_grouped, 0, col_num],  # Name from header row
            'categories': [sheet_grouped, 1, 0, len(pivot_df), 0],  # Categories (ReportMonth)
            'values':     [sheet_grouped, 1, col_num, len(pivot_df), col_num], 
            'data_labels': {'value': True}
        })

    # Configure chart title and axes
    chart.set_title({'name': f'VIRP Violations for {client_name}'})
    chart.set_x_axis({'name': 'Report Month'})
    chart.set_y_axis({'name': 'Distinct Cases'})
    chart.set_legend({'position': 'bottom'})

    # Insert the chart into the Grouped Data sheet
    worksheet_grouped.insert_chart('H2', chart)

    # Save and close the workbook
    workbook.close()

# Example usage:
create_dynamic_bar_chart('./Master_Data.xlsx', 'output_chart_with_grouped_data1.xlsx')
