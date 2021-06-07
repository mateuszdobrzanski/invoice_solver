from fakturownia import get_last_12m_invoices
from functions import return_xls_sheet, return_xls_row, return_customized_xls_header, return_dict_from_lists, \
    return_split_dist, check_tax_numbers, return_invoice_no

# xls_file = 'C:\\invoice.xls'
xls_file = 'C:\\Users\\Mateusz\\Desktop\\aaa.xls'
# open xls file
sheet = return_xls_sheet(xls_file)

# get customized xls header
xls_header = return_customized_xls_header(return_xls_row(sheet, 0))

# iterate over the xls lines
# [1:] skipping header
for x in range(sheet.nrows)[1:]:
    row = return_xls_row(sheet, x)

    # convert row to dictionary
    values_dict = return_dict_from_lists(xls_header, row)

    source_dict = return_split_dist(values_dict)

    if check_tax_numbers(source_dict)['status'] == 'warning':
        source_dict['NIP'] = check_tax_numbers(source_dict)['new_val']

    # print(source_dict)
    #
    invoice_number = return_invoice_no(source_dict)

    if invoice_number['status'] == 'success':
        print(source_dict)
        tax_id = source_dict['NIP']
        print(invoice_number['val'])

        # json_output = get_last_12m_invoices(tax_id, xls_file)
        json_output = {'status': 'success',
                       'message': 'correctly downloaded invoices',
                       'val': 'C:\\invoices\\test.json'}
        print(json_output)


    # print('\n')

