from functions import return_xls_sheet, return_xls_row, return_customized_xls_header, return_dict_from_lists, \
    return_split_dist, check_tax_numbers, return_invoice_no

xls_file = 'C:\\invoice.xls'

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

    print(source_dict)
    # print(check_tax_numbers(source_dict))
    #
    print(return_invoice_no(source_dict))


    print('\n')

