import os
import re
import json
import xlrd

from fakturownia import change_invoice_status_to_paid, change_invoice_status_to_partial
from datetime import datetime


def return_xls_workbook(xls_file_path):
    workbook = xlrd.open_workbook(xls_file_path)
    return workbook


def return_xls_sheet(workbook):
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


def check_tax_numbers(dictionary):
    if 'PAYPRO' in dictionary['Nazwa i adres Kontrahenta']:
        status = {'status': 'error',
                  'message': 'PAYPRO payemnt'}
        return status

    if 'Na rachunek wirtualny' in dictionary:
        virtual_bill = dictionary['Na rachunek wirtualny']
        virtual_bill = virtual_bill.replace(" ", "")
        virtual_bill = virtual_bill[-10:]
    else:
        status = {'status': 'error',
                  'val': False,
                  'message': 'key not found'}
        return status

    if dictionary['NIP'] == virtual_bill:
        status = {'status': 'success',
                  'val': True,
                  'message': ''}
        return status
    else:
        status = {'status': 'warning',
                  'val': False,
                  'new_val': virtual_bill,
                  'message': 'tax numbers are different'}
        return status


def remove_delimiters(text):
    delimiters = [" ", ",", ".", "!", "?", "/", "\\", "&", "-", "_", ":", ";", "@", "'", "..."]

    for d in delimiters:
        text = text.replace(d, '')

    return text


# we are trying to find invoice number or invoice number with prefix
# also, we are cleaning other chars
# when we meet multiple numbers in one title, we skip this row
def find_number_by_re(text):
    # we set that max and min length value of our number
    min_length = 17
    max_length = 28
    result = None

    for length in range(max_length, min_length, -1):
        s = '([0-9]{' + str(length) + '})'
        pattern = re.compile(s)
        result = pattern.search(text)
        if result is not None:
            if len(re.findall(pattern, text)) == 1:
                p = '(FAB|FSA|FWI|FUS)([0-9]{' + str(length) + '})'
                pattern = re.compile(p)
                r = pattern.search(text)
                if r is not None:
                    return r.group(0)
                else:
                    return result.group(0)
            else:
                return 'error - to many cases'
    if result is None:
        return 'error - no prefix and number in title'


# return cleaned up invoice number
def return_invoice_no(dictionary):
    cleanup_title = ''
    cleanup_sec_title = ''

    if 'Tytuł' in dictionary:
        # raw_title = dictionary['Tytuł']
        cleanup_title = find_number_by_re(remove_delimiters(dictionary['Tytuł']))
        # print('Tytuł: ' + raw_title + " | " + cleanup_title)

    if 'Numer faktury' in dictionary:
        # raw_sec_title = dictionary['Numer faktury']
        cleanup_sec_title = find_number_by_re(remove_delimiters(dictionary['Numer faktury']))
        # print('Numer faktury: ' + raw_sec_title + " | " + cleanup_sec_title)

    if cleanup_sec_title != '':
        if 'error' in cleanup_sec_title:
            status = {'status': 'error',
                      'message': 'error occurred in "Numer faktury"',
                      'message_detail': cleanup_sec_title}
        else:
            if cleanup_title != '' and not 'error' in cleanup_title:
                if cleanup_title == cleanup_sec_title:
                    status = {'status': 'success',
                              'message': 'both values are same',
                              'val': cleanup_title}

                elif cleanup_title in cleanup_sec_title:
                    status = {'status': 'success',
                              'message': '"Tytuł" is in "Numer faktury" ',
                              'val': cleanup_sec_title}

                elif cleanup_sec_title in cleanup_title:
                    status = {'status': 'success',
                              'message': '"Numer faktury" is in "Tytuł"',
                              'val': cleanup_title}
                else:
                    status = {'status': 'error',
                              'message': 'both values are different',
                              'message_detail': cleanup_title + " " + cleanup_sec_title}
            else:
                status = {'status': 'success',
                          'message': 'returned value from "Numer faktury"',
                          'val': cleanup_sec_title}
    else:
        if 'error' in cleanup_title:
            status = {'status': 'error',
                      'message': 'error occurred in "Tytuł"',
                      'message_detail': cleanup_title}
        else:
            status = {'status': 'success',
                      'message': 'returned value from "Tytuł"',
                      'val': cleanup_title}

    return status


def open_json_file(file_name):
    with open(file_name, encoding='utf-8', errors='ignore') as json_data:
        f = json.load(json_data)

    return f


def return_invoice(json_file, xls_invoice_number):
    j_file = open_json_file(json_file)
    status = {'status': 'error',
              'message': 'invoice number was not found'}

    for j in j_file:
        cleaned_number = remove_delimiters(j['number'])

        if cleaned_number == xls_invoice_number:
            status = {'status': 'success',
                      'message': 'founded properly invoice number',
                      'val': j,
                      'message_detail': j['number'] + " | " + cleaned_number + " | " + xls_invoice_number}
            return status
        elif xls_invoice_number in cleaned_number:
            status = {'status': 'success',
                      'message': 'founded number from xls in invoice number',
                      'val': j,
                      'message_detail': j['number'] + " | " + cleaned_number + " | " + xls_invoice_number}
            return status
        else:
            status = {'status': 'error',
                      'message': 'invoice number was not found',
                      'message_detail': j['number'] + " | " + cleaned_number + " | " + xls_invoice_number}
    return status


def compare_amounts(json_data, amount_xls):
    amount_json = float(json_data['val']['price_gross'])

    if amount_json == float(amount_xls):
        returned_status = change_invoice_status_to_paid(json_data['val']['id'])
        status = {'status': 'success',
                  'val': 'paid',
                  'message': 'changed to paid',
                  'message_detail': returned_status}
    elif amount_json > float(amount_xls):
        returned_status = change_invoice_status_to_partial(json_data['val']['id'], amount_xls)
        status = {'status': 'success',
                  'val': 'partial',
                  'message': 'changed to partial',
                  'message_detail': returned_status}
    elif amount_json < float(amount_xls):
        # TODO change status to overpaid
        status = {'status': 'success',
                  'val': 'partial',
                  'message': 'should be changed to overpaid'}

    return status


def compare_json_xls(json_data, xls_data):
    if json_data['val']['status'] == 'issued':
        status = compare_amounts(json_data, xls_data['Kwota'])

    elif json_data['val']['status'] == 'paid':
        status = {'status': 'success',
                  'val': 'paid',
                  'message': 'status paid'}

    elif json_data['val']['status'] == 'partial':
        json_paid = float(json_data['val']['paid'])
        xls_paid = float(xls_data['Kwota'])

        amount = json_paid + xls_paid

        status = compare_amounts(json_data, amount)

    elif json_data['val']['status'] == 'sent':
        status = {'status': 'warning',
                  'val': 'sent',
                  'message': 'status - sent'}

    elif json_data['val']['status'] == 'rejected':
        status = {'status': 'warning',
                  'val': 'rejected',
                  'message': 'status - rejected'}

    else:
        status = {'status': 'error',
                  'message': 'error occurred'}

    return status


def output_filename(file_path, timestamp):
    head, tail = os.path.split(file_path)

    return head + "\\" + timestamp + "-" + tail


def return_date_time():
    dt_now = datetime.now()

    return dt_now.strftime("%d%m%Y-%H%M%S")

