from functions import return_xls_sheet, return_xls_row, return_customized_xls_header, return_dict_from_lists, \
    return_split_dist

xls_file = 'C:\\invoice.xls'
sheet = return_xls_sheet(xls_file)

xls_header = return_customized_xls_header(return_xls_row(sheet, 0))

# iterate over the xls lines
# [1:] skipping header
for x in range(sheet.nrows)[1:]:
    row = return_xls_row(sheet, x)

    values_dict = return_dict_from_lists(xls_header, row)
    # print(str(n_dict))

    print(str(return_split_dist(values_dict)))

