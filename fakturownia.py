from pathlib import Path
import os

import requests
import json
import codecs
import configparser


config = configparser.ConfigParser()
config.read('settings/settings.ini')

API_TOKEN = config['DEFAULT']['api_token']
API_URL = config['DEFAULT']['api_url']

API_URL_BASE = API_URL[:-1] + '.json'
API_URL_CLIENTS = config['DEFAULT']['api_url_clients']

HEADERS = {'Accept: application/json',
           'Content-Type: application/json'}


def is_customer_exist(customer_tax_id):
    params = (
        ('api_token', API_TOKEN),
        ('tax_no', customer_tax_id),
    )

    response = requests.get(API_URL_CLIENTS, params=params)
    temp = response.text

    if response.status_code == 200 and len(temp) > 2:
        status = {'status': 'success',
                  'message': 'customer exist',
                  'val': True}
    else:
        status = {'status': 'success',
                  'message': 'customer exist',
                  'val': True}

    return status


# download invoices from the last 12 months using customer tax id
# invoices saved where is source file
# returned file name where we had all invoices
def get_last_12m_invoices(customer_tax_id, file_path):
    page_no_data = 1
    my_data = ""
    not_empty_data = True
    output_filename = ''

    # download invoices while we have data
    if is_customer_exist(customer_tax_id)['status'] == 'success':
        while not_empty_data is True:
            url_string = API_URL_BASE + "?order=&income=yes&query=" + \
                         customer_tax_id + "&kind=all&period=last_12_months&search_date_type=issue_date&api_token=" \
                         + API_TOKEN + "&page=" + str(page_no_data)

            response = requests.get(url_string)
            if len(response.text) > 2:
                json_output = response.text
                parsed = json.loads(json_output)
                my_data = my_data + json.dumps(parsed, indent=4, sort_keys=True, ensure_ascii=False)
                page_no_data = page_no_data + 1
            else:
                not_empty_data = False

        # here, we save the downloaded invoices
        head, tail = os.path.split(file_path)
        path = head + "\\" + "invoices\\"

        # create directory when not exist
        Path(path).mkdir(parents=True, exist_ok=True)

        output_filename = path + customer_tax_id + ".json"

        if '][' in my_data:
            my_data = my_data.replace('][', ',')

        # save file with utf-8 encoding
        with codecs.open(output_filename, 'w', "utf-8") as outfile:
            outfile.write(my_data)

        status = {'status': 'success',
                  'message': 'correctly downloaded invoices',
                  'val': output_filename}
    else:
        status = {'status': 'error',
                  'message': 'an error has occurred when downloaded invoices'}

    return status
