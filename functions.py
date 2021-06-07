import xlrd


def return_xls_sheet(xls_file_path):
    workbook = xlrd.open_workbook(xls_file_path)
    return workbook.sheet_by_index(0)


def return_xls_row(workbook_sheet, row_number):
    return workbook_sheet.row_values(row_number)


# source header has not all tax number column, so we need to add a custom description
def return_customized_xls_header(xls_header):
    return ['NIP' if header_cell == '' else header_cell for header_cell in xls_header]


def return_dict_from_lists(key, value):
    return {key[i]: value[i] for i in range(len(key))}


def remove_key(dictionary, key):
    new_dict = dict(dictionary)
    del new_dict[key]
    return new_dict


# 'Dane operacji' (operation data) - is really dictionary, to get the data easier at first we get values using this key,
# next, we create a dictionary (splitting by characters like ":" and "|") and combine two dictionaries
def return_split_dist(dictionary):
    op_data = dictionary['Dane operacji']

    transfer_list = op_data.split('|')
    transfer_list = list(filter(None, transfer_list))

    # get first delimiter(:) to create key value pair based on list
    transfer_dict = dict(element.split(': ', 1) for element in transfer_list)

    # delete values without key/value pair
    new_dict = remove_key(dictionary, 'Dane operacji')

    # combine two dict to one
    return {**new_dict, **transfer_dict}
