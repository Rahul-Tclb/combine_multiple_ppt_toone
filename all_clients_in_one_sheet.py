import pandas as pd
import xlsxwriter


def create_combined_sheet_with_charts(filename, output_excel):
    """
    Writes original data, grouped data, and charts for multiple clients in a single Excel sheet.
    Each client's data is written consecutively with a gap in between.

    Args:
        filename: The name of the input Excel file.
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

    bold_format = workbook.add_format({'bold': True , 'border': 1 } )
    border_format = workbook.add_format({'border': 1})
    

    # Row offset to manage spacing between sections
    row_offset = 0

    # Define color palette
    colors = [
        {'color': '#1434CB'},   # Visa Blue: R, G, B = 20, 52, 203
        {'color': '#FCC015'},  # Visa Yellow : R, G, B = 252,192,21,
        {'color': '#E1E1E1'},  # White: R, G, B = 225, 225, 225
        {'color': '#C0BDBB'}, # Light Gray
        # {'color': '#000000'},  # Black: R, G, B = 0, 0, 0
        # {'color': '#021E4C'}, # Visa Dark 2,30,76
        {'color': '#96918C'},  # Gray: R, G, B = 150, 145, 140
        {'color': '#F0F0F0'}  # Off-white: R, G, B = 240, 240, 240  
    ]

    for client_name in clients:
        # Filter data for the client
        client_df = df[df['ClientName'] == client_name]

        # Group the data by 'ReportMonth' and 'Status' and sum 'DistinctCases'
        grouped_df = client_df.groupby(['ReportMonth', 'Status']).sum().reset_index()

        # Create a pivot table
        pivot_df = grouped_df.pivot_table(index='ReportMonth', columns='Status', values='DistinctCases', fill_value= "")
        pivot_df.reset_index(inplace=True)

        # Write the client name as a header
        worksheet.write(row_offset, 0, f"Client: {client_name}" , bold_format )

        # Write original data
        original_header = list(client_df.columns)
        worksheet.write_row(row_offset + 2, 0, original_header , bold_format )  # Write headers
        for i, row in enumerate(client_df.values, start=row_offset + 3):
            worksheet.write_row(i, 0, row , border_format)


        # Set column width to 32 for all columns
        for col_num in range(len(original_header)):
            worksheet.set_column(col_num, col_num, 32)

        # Write grouped data below original data
        grouped_start_row = row_offset + 3 + len(client_df) + 2  # Leave 2 rows gap
        worksheet.write_row(grouped_start_row, 0, ['ReportMonth'] + list(pivot_df.columns[1:]) , bold_format )  # Write headers
        for i, row in enumerate(pivot_df.values, start=grouped_start_row + 1):
            worksheet.write_row(i, 0, row , border_format)
        
        # Set column width to 32 for grouped data
        for col_num in range(len(['ReportMonth'] + list(pivot_df.columns[1:]))):
            worksheet.set_column(col_num, col_num, 32)

        # Create a stacked column chart
        chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

        # Add data series dynamically with specified colors
        for col_num in range(1, pivot_df.shape[1]):
            chart.add_series({
                'name':       ['Client Data and Charts', grouped_start_row, col_num],
                'categories': ['Client Data and Charts', grouped_start_row + 1, 0, grouped_start_row + len(pivot_df), 0],
                'values':     ['Client Data and Charts', grouped_start_row + 1, col_num, grouped_start_row + len(pivot_df), col_num],
                'data_labels': {'value': True},
                'fill': colors[(col_num - 1) % len(colors)]  # Cycle through the defined color palette
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
create_combined_sheet_with_charts('./Master_Data.xlsx', '1_combined_client_data_and_charts.xlsx')
