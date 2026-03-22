import xlsxwriter

def create_sample_excel():
    workbook = xlsxwriter.Workbook('sample_format.xlsx')
    worksheet = workbook.add_worksheet('Transactions')

    # Formatting
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
    money_format = workbook.add_format({'num_format': '$#,##0.00'})

    # Write Headers
    headers = ['Date', 'Description', 'Amount', 'Category (Optional)']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    # Write Data
    data = [
        ['10/01/2023', 'Netflix Subscription', 15.99, 'Entertainment'],
        ['10/02/2023', 'Uber Rides', 24.50, 'Transportation'],
        ['10/04/2023', 'Whole Foods Market', 112.30, 'Groceries'],
        ['10/05/2023', 'Amazon Prime Video', 8.99, 'Entertainment'],
        ['10/08/2023', 'Spotify Premium', 10.99, 'Entertainment'],
        ['10/12/2023', 'Planet Fitness', 29.99, 'Health & Fitness'],
        ['10/15/2023', 'Uber Eats', 35.50, 'Dining Out'],
        ['10/18/2023', 'Adobe Creative Cloud', 54.99, 'Software'],
    ]

    for row_num, row_data in enumerate(data, start=1):
        worksheet.write(row_num, 0, row_data[0], date_format)
        worksheet.write(row_num, 1, row_data[1])
        worksheet.write(row_num, 2, row_data[2], money_format)
        worksheet.write(row_num, 3, row_data[3])

    # Adjust column widths
    worksheet.set_column('A:A', 12)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 12)
    worksheet.set_column('D:D', 20)

    workbook.close()
    print("Excel template created cleanly.")

if __name__ == '__main__':
    create_sample_excel()
