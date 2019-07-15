import re
from statements import plq_id_plq_qty_per, orl_id_orl_qty_ready


# Split payload string, return named variables
def payload_handler(raw_payload):
    sigm_string = raw_payload.split(', ')[-1]
    payload = {
        'db_ref': raw_payload.split(', ')[0],
        'db_ref_type': raw_payload.split(', ')[1],
        'label_type': raw_payload.split(', ')[2],
        'qty_ref': raw_payload.split(', ')[3],
        'label_name': raw_payload.split(', ')[4] if raw_payload.split(', ')[2] == 'GENERIC' else '',
        'station': re.findall(r'(?<= w)(.*)$', sigm_string)[0]
    }
    return payload


# Generate range of serial numbers from 'XXXXX to XXXXX' formatted string
def sn_range_to_list(raw_sn_range, sn_list):
    first_sn = raw_sn_range[0:5]
    last_sn = raw_sn_range[-5:]
    sn_range = range(int(first_sn), int(last_sn) + 1)

    for sn in sn_range:
        sn_list.append(sn)
    return sn_list


# Return list of serial numbers from string stored in planning book note field (planning_lot_quantity.plq_note)
def serial_no_range(plq_note):
    # Allows users to enter comments after serial numbers
    serial_numbers = plq_note.split(';')
    serial_numbers = serial_numbers[0].strip('SN: ')

    sn_list = []
    if ', ' in serial_numbers:
        serial_numbers = serial_numbers.split(', ')
        for item in serial_numbers:
            if ' to ' in item:
                raw_sn_range = item
                sn_list = sn_range_to_list(raw_sn_range, sn_list)
            else:
                serial_number = item
                sn_list.append(serial_number)
    elif ' to ' in serial_numbers:
        raw_sn_range = serial_numbers
        sn_list = sn_range_to_list(raw_sn_range, sn_list)
    else:
        serial_number = serial_numbers
        sn_list.append(serial_number)

    return sn_list


def modulo_43(string):
    ref = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-._$/+%'
    total = 0
    for character in string:
        total += ref.find(character)
    remainder = total % 43
    return ref[remainder]


# Return label quantity based on production/order line quantity, default to 1
def select_print_qty(config, payload):
    qty = 1
    if payload['qty_ref'] == 'MULTI':
        if payload['db_ref_type'] == 'plq_id':
            qty = plq_id_plq_qty_per(config, payload['plq_id'])

        elif payload['db_ref_type'] == 'orl_id':
            qty = orl_id_orl_qty_ready(config, payload['orl_id'])
    return qty


def get_cli_no_customer(config, payload):
    if payload['label_type'] == 'GENERIC':
        return 'generic'
    for customer in config.CUSTOMERS:
        if payload['cli_no'] in config.CUSTOMERS[customer]:
            return customer
    return 'generic'
