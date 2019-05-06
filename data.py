from sql import plq_id_plq_note, plq_id_plq_qty_per, orl_id_orl_qty


# Generate range of serial numbers from 'XXXXX to XXXXX' formatted string
def sn_range_to_list(raw_sn_range, sn_list):
    first_sn = raw_sn_range[0:5]
    last_sn = raw_sn_range[-5:]
    sn_range = range(int(first_sn), int(last_sn) + 1)

    for sn in sn_range:
        sn_list.append(sn)
    return sn_list


# Return list of serial numbers from string stored in planning book note field (planning_lot_quantity.plq_note)
def serial_no_range(plq_id):
    plq_note = plq_id_plq_note(plq_id)
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


# Return label quantity based on production/order line quantity, default to 1
def select_print_qty(qty_ref, db_ref_type, ref):
    qty = 1
    if qty_ref == 'MULTI':
        if db_ref_type == 'plq_id':
            plq_id = ref
            qty = plq_id_plq_qty_per(plq_id)
        elif db_ref_type == 'orl_id':
            orl_id = ref
            qty = orl_id_orl_qty(orl_id)

    return qty


# if __name__ == "__main__":
#     sn_range_to_list('12345 to 67890', [])
