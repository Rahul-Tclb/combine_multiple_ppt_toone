import pandas as pd
import xlsxwriter


def create_combined_sheet_with_charts(filename, output_excel):
    """
    Writes original data, grouped data, and charts for multiple clients in a single Excel sheet.
    Each client's data is written consecutively with a gap in between.

    Args:
        filename: The name of the input excel file.
        output_excel: The name of the output Excel file.
    """
    # Load the data
    df = pd.read_excel(filename)
    
    # Get unique client names
    clients = df['ClientName'].unique()

    # Create the Excel workbook
    workbook = xlsxwriter.Workbook(output_excel)

    # Add a single worksheet
    worksheet = workbook.add_worksheet("Client Data and Charts")

    # Row offset to manage spacing between sections
    row_offset = 0

    for client_name in clients:
        # Filter data for the client
        client_df = df[df['ClientName'] == client_name]

        # Group the data by 'ReportMonth' and 'Status' and sum 'DistinctCases'
        grouped_df = client_df.groupby(['ReportMonth', 'Status']).sum().reset_index()

        # Create a pivot table
        pivot_df = grouped_df.pivot_table(index='ReportMonth', columns='Status', values='DistinctCases', fill_value=0)
        pivot_df.reset_index(inplace=True)

        # Replace zeros with blanks
        pivot_df = pivot_df.replace(0, "")

        # Write the client name as a header
        worksheet.write(row_offset, 0, f"Client: {client_name}")

        # Write original data
        original_header = list(client_df.columns)
        worksheet.write_row(row_offset + 2, 0, original_header)  # Write headers
        for i, row in enumerate(client_df.values, start=row_offset + 3):
            worksheet.write_row(i, 0, row)

        # Write grouped data below original data
        grouped_start_row = row_offset + 3 + len(client_df) + 2  # Leave 2 rows gap
        worksheet.write_row(grouped_start_row, 0, ['ReportMonth'] + list(pivot_df.columns[1:]))  # Write headers
        for i, row in enumerate(pivot_df.values, start=grouped_start_row + 1):
            worksheet.write_row(i, 0, row)

        # Create a clustered column chart
        chart = workbook.add_chart({'type': 'column' , 'subtype':'stacked'})

        # Add data series dynamically
        for col_num in range(1, pivot_df.shape[1]):
            chart.add_series({
                'name':       ['Client Data and Charts', grouped_start_row, col_num],
                'categories': ['Client Data and Charts', grouped_start_row + 1, 0, grouped_start_row + len(pivot_df), 0],
                'values':     ['Client Data and Charts', grouped_start_row + 1, col_num, grouped_start_row + len(pivot_df), col_num],
                'data_labels': {'value': True }
            })

        # Configure the chart
        chart.set_title({'name': f'VIRP Violations for {client_name}'})
        chart.set_x_axis({'name': 'Report Month', 'major_gridlines': {'visible': False}})
        chart.set_y_axis({'name': 'Distinct Cases', 'major_gridlines': {'visible': False}})
        chart.set_legend({'position': 'bottom'})
        chart.set_size({'width': 900, 'height': 500})

        # Insert the chart below grouped data
        chart_start_row = grouped_start_row + len(pivot_df) + 2  # Leave 2 rows gap
        worksheet.insert_chart(chart_start_row, 0, chart)

        # Update row offset for the next client
        row_offset = chart_start_row + 30  # Add padding for the next client's data

    # Save and close the workbook
    workbook.close()


# Example usage:
create_combined_sheet_with_charts('./Master_Data.xlsx', 'combined_client_data_and_charts.xlsx')
