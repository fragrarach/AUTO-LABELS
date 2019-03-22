import psycopg2.extensions
import re
import os
from win32com.client import Dispatch
from shutil import copyfile

# PostgreSQL DB connection configs
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)


# Check whether app should reference dev or prod server/db
def dev_check():
    raw_filename = os.path.basename(__file__)
    removed_extension = raw_filename.split('.')[0]
    last_word = removed_extension.split('_')[-1]
    if last_word == 'DEV':
        return True
    else:
        return False


# Initialize production DB connection, listen cursor and query cursor
def sigm_conn():
    global conn_sigm, sigm_query
    if dev_check():
        conn_sigm = psycopg2.connect("host='192.168.0.57' dbname='DEV' user='SIGM' port='5493'")
    else:
        conn_sigm = psycopg2.connect("host='192.168.0.250' dbname='QuatroAir' user='SIGM' port='5493'")
    conn_sigm.set_client_encoding("latin1")
    conn_sigm.set_isolation_level(psycopg2.extensions.ISOLATION_LEVEL_AUTOCOMMIT)

    sigm_listen = conn_sigm.cursor()
    sigm_listen.execute("LISTEN labels;")
    sigm_query = conn_sigm.cursor()

    return conn_sigm, sigm_query


# Dymo COM configs
labelCom = Dispatch('Dymo.DymoAddIn')
labelText = Dispatch('Dymo.DymoLabels')


# TODO : Create module for tabular/scalar data and production query functions so they can be imported/reused

# Convert tabular query result to list (2D array)
def tabular_data(result_set):
    lines = []
    for row in result_set:
        line = []
        for cell in row:
            if type(cell) == str:
                cell = cell.strip()
            line.append(cell)
        lines.append(line)
    return lines


# Convert scalar query result to singleton variable of any data type
def scalar_data(result_set):
    for row in result_set:
        for cell in row:
            if type(cell) == str:
                cell = cell.strip()
            return cell


# Query production database
def production_query(sql_exp):
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()
    return result_set


# TODO : Implement tabular/scalar data and production query functions in SQL functions

# Pull 'prt_no' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_prt_no(plq_id):
    sql_exp = f'SELECT trim(prt_no) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            prt_no = cell
            return prt_no


# Pull 'prt_desc1' record from 'part' table based on 'prt_no' record
def prt_no_prt_desc(prt_no):
    sql_exp = f'SELECT trim(prt_desc1) FROM part WHERE prt_no = \'{prt_no}\''
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            prt_desc1 = cell
            return prt_desc1


# Pull 'plq_qty_per' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_plq_qty_per(plq_id):
    sql_exp = f'SELECT plq_qty_per FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            plq_qty_per = cell
            return plq_qty_per


# Pull 'plq_qty_per' record from 'planning_lot_quantity' table based on 'plq_id' record
def plq_id_plq_note(plq_id):
    sql_exp = f'SELECT trim(plq_note) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            plq_note = cell
            return plq_note


# Pull 'orl_id' record from 'order_line' table based on 'ord_no' record
def ord_no_orl_id(ord_no):
    sql_exp = f'SELECT orl_id FROM order_line WHERE ord_no = {ord_no} AND prt_no <> \'\''
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    orl_ids = []
    for row in result_set:
        for cell in row:
            orl_id = cell
            orl_ids.append(orl_id)
    return orl_ids


# Pull 'orl_quantity' record from 'order_line' table based on 'orl_id' record
def orl_id_orl_qty(orl_id):
    sql_exp = f'SELECT (orl_quantity)::INT FROM order_line WHERE orl_id = {orl_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            orl_quantity = cell
            return orl_quantity


# Pull 'prt_no' record from 'order_line' table based on 'orl_id' record
def orl_id_prt_no(orl_id):
    sql_exp = f'SELECT trim(prt_no) FROM order_line WHERE orl_id = {orl_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            prt_no = cell
            return prt_no


# Pull 'prt_no' record from 'order_line' table based on 'orl_id' record
def orl_id_prt_desc(orl_id):
    sql_exp = f'SELECT trim(prt_desc) FROM order_line WHERE orl_id = {orl_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            prt_desc = cell
            return prt_desc


# Return list of serial numbers from string stored in planning book note field (planning_lot_quantity.plq_note)
def serial_no_range(plq_id):
    plq_note = plq_id_plq_note(plq_id)
    # Allows users to enter comments after serial numbers
    serial_numbers = plq_note.split(';')
    serial_numbers = serial_numbers[0].strip('SN: ')

    sn_list = []

    # Generate range of serial numbers from 'XXXXX to XXXXX' formatted string
    def sn_range_to_list(raw_sn_range):
        first_sn = raw_sn_range[0:5]
        last_sn = raw_sn_range[-5:]
        sn_range = range(int(first_sn), int(last_sn) + 1)
        for sn in sn_range:
            sn_list.append(sn)

    if ', ' in serial_numbers:
        serial_numbers = serial_numbers.split(', ')
        for item in serial_numbers:
            if ' to ' in item:
                raw_sn_range = item
                sn_range_to_list(raw_sn_range)
            else:
                serial_number = item
                sn_list.append(serial_number)
    elif ' to ' in serial_numbers:
        raw_sn_range = serial_numbers
        sn_range_to_list(raw_sn_range)
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


# Pull list of currently online shared Dymo printers on the network
def get_dymo_printers():
    printers = labelCom.GetDymoPrinters()
    print(printers)


# Return printer based on which computer is passing the payload
def select_printer(label_ref, station):
    control_panel_printer = r'DEV DYMO'
    unit_printer = r'DEV DYMO'
    
    if station == 'DESKTOP-PKO4ES1':
        shipping_printer = f'\\\\DESKTOP-PKO4ES1\\ERICK DYMO'
    elif station == 'CAD2':
        shipping_printer = f'\\\\CAD2.aerofil.local\\DEV SHIPPING DYMO'
    elif station == 'DESKTOP-FG8A5QJ':
        shipping_printer = f'\\\\DESKTOP-FG8A5QJ.aerofil.local\\GATTA DYMO'
    elif station == 'DESKTOP-69UD4SH':
        shipping_printer = f'\\\\DESKTOP-69UD4SH\\JORGE DYMO'
    else:
        shipping_printer = f'\\\\DESKTOP-FG8A5QJ.aerofil.local\\GATTA DYMO'

    serial_number_printer = f'\\\\DESKTOP-PKO4ES1.aerofil.local\\DYMO 450 SERIAL'

    if label_ref == 'CONTROL PANEL':
        return control_panel_printer
    elif label_ref == 'UNIT':
        return unit_printer
    elif label_ref == 'SHIPPING':
        return shipping_printer
    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return shipping_printer
    elif label_ref == 'SERIAL NUMBER':
        return serial_number_printer


# Return label based on which report is being run
def select_label(label_ref):
    label_dir = \
        r'E:\DATA\Fortune\SIGMWIN.DTA\QuatroAir\Documents\REFERENCE FILES\AUTOMATED LABEL GENERATION\DYMO LABELS'

    control_panel_label = label_dir + '\CONTROL PANEL.label'
    unit_label = label_dir + r'\UNIT.label'
    shipping_label = label_dir + r'\SHIPPING.label'
    shipping_serial_label = label_dir + r'\SHIPPING SERIAL NUMBER.label'
    serial_number_label = label_dir + r'\SERIAL NUMBER.label'

    if label_ref == 'CONTROL PANEL':
        return control_panel_label
    elif label_ref == 'UNIT':
        return unit_label
    elif label_ref == 'SHIPPING':
        return shipping_label
    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return shipping_serial_label
    elif label_ref == 'SERIAL NUMBER':
        return serial_number_label


# Create a working copy of label template
def copy_label_template(label):
    label_name = label.split('\\')[-1]
    label_dir = '\\'.join(label.split('\\')[:-1])
    label_template = f'{label_dir}\\TEMPLATES\\{label_name}'
    copyfile(label_template, label)


# Edit label text based on label/DB reference passed by payload
def label_text_handler(db_ref_type, db_ref, label_ref, printer, qty):
    # TODO : Put 'select_label' function in the main loop, pass label instead of label_ref
    label = select_label(label_ref)
    copy_label_template(label)

    if db_ref_type == 'plq_id':
        # PLETI Report
        plq_id = db_ref
        prt_no = plq_id_prt_no(plq_id)
        prt_desc = prt_no_prt_desc(prt_no)
        serial_no_list = serial_no_range(plq_id)

        if label_ref == 'CONTROL PANEL':
            pairs = []
            prt_no_pair = ['<Text></Text>', f'<Text>{prt_no}</Text>']
            pairs.append(prt_no_pair)

            print_label(pairs, label, printer, qty)

        elif label_ref == 'UNIT':
            for serial_no in serial_no_list:
                pairs = []
                serial_no = str(serial_no)
                serial_no_pair = ['<Text>1</Text>', f'<Text>{serial_no}</Text>']
                pairs.append(serial_no_pair)
                prt_no_pair = ['<Text>2</Text>', f'<Text>{prt_no}</Text>']
                pairs.append(prt_no_pair)

                print_label(pairs, label, printer)

        elif label_ref == 'SERIAL NUMBER':
            for serial_no in serial_no_list:
                pairs = []
                serial_no = str(serial_no)
                serial_no_pair = ['<Text></Text>', f'<Text>{serial_no}</Text>']
                pairs.append(serial_no_pair)

                print_label(pairs, label, printer, qty)

        elif label_ref == 'SHIPPING SERIAL NUMBER':
            for serial_no in serial_no_list:
                pairs = []
                serial_no = str(serial_no)
                prt_no_barcode_pair = ['<Text>3</Text>', f'<Text>{prt_no}</Text>']
                pairs.append(prt_no_barcode_pair)
                serial_no_barcode_pair = ['<Text>12345</Text>', f'<Text>{serial_no}</Text>']
                pairs.append(serial_no_barcode_pair)
                prt_no_pair = ['<String xml:space="preserve">: 3</String>',
                               f'<String xml:space="preserve">: {prt_no}</String>']
                pairs.append(prt_no_pair)
                prt_desc_pair = ['<String xml:space="preserve">: 2</String>',
                                 f'<String xml:space="preserve">: {prt_desc}</String>']
                pairs.append(prt_desc_pair)
                serial_no_pair = ['<String xml:space="preserve">: 12345</String>',
                                  f'<String xml:space="preserve">: {serial_no}</String>']
                pairs.append(serial_no_pair)

                print_label(pairs, label, printer, qty)

    elif db_ref_type == 'orl_id':
        # CCETI Report
        orl_id = db_ref
        prt_no = orl_id_prt_no(orl_id)
        prt_desc = orl_id_prt_desc(orl_id)
        if label_ref == 'SHIPPING':
            pairs = []
            prt_no_barcode_pair = ['<Text>3</Text>', f'<Text>{prt_no}</Text>']
            pairs.append(prt_no_barcode_pair)
            prt_no_pair = ['<String xml:space="preserve">: 3</String>',
                           f'<String xml:space="preserve">: {prt_no}</String>']
            pairs.append(prt_no_pair)
            prt_desc_pair = ['<String xml:space="preserve">: 2</String>',
                             f'<String xml:space="preserve">: {prt_desc}</String>']
            pairs.append(prt_desc_pair)

            print_label(pairs, label, printer, qty)


# Set text label text, print label, revert label text
def print_label(pairs, label, printer, qty=1):
    insert_label_text(pairs, label)
    dymo_print(printer, label, qty)
    reset_label_text(pairs, label)


# Dymo COM API functions
def dymo_print(printer, label, print_qty=1):
    labelCom.SelectPrinter(printer)
    labelCom.Open(label)
    labelCom.StartPrintJob()
    labelCom.Print(print_qty, False)
    labelCom.EndPrintJob()


# Pass list of pairs of strings (old text, new text) to edit label XML text
def insert_label_text(pairs, label):
    for pair in pairs:
        with open(label, "r") as file:
            file_data = file.read()
        file_data = file_data.replace(pair[0], pair[1])

        with open(label, "w") as file:
            file.write(file_data)


# TODO : Test if copying template is faster/more consistent than reverting text changes
# Pass list of pairs of strings (old text, new text) to revert label XML text
def reset_label_text(pairs, label):
    for pair in pairs:
        with open(label, "r") as file:
            file_data = file.read()
        file_data = file_data.replace(pair[1], pair[0])

        with open(label, "w") as file:
            file.write(file_data)


# Split payload string, return named variables
def payload_handler(payload):
    db_ref = payload.split(', ')[0]
    db_ref_type = payload.split(', ')[1]
    label_ref = payload.split(', ')[2]
    qty_ref = payload.split(', ')[3]
    sigm_string = payload.split(', ')[4]
    station = re.findall(r'(?<= w)(.*)$', sigm_string)[0]

    return db_ref, db_ref_type, label_ref, qty_ref, station


def main():
    global conn_sigm, sigm_query
    conn_sigm, sigm_query = sigm_conn()
    get_dymo_printers()
    while 1:
        try:
            conn_sigm.poll()
        except:
            print('Database cannot be accessed, PostgreSQL service probably rebooting')
            try:
                conn_sigm.close()
                conn_sigm, sigm_query = sigm_conn()

            except:
                pass
            else:
                get_dymo_printers()
        else:
            conn_sigm.commit()
            while conn_sigm.notifies:
                notify = conn_sigm.notifies.pop()
                raw_payload = notify.payload

                db_ref, db_ref_type, label_ref, qty_ref, station = payload_handler(raw_payload)

                printer = select_printer(label_ref, station)

                if db_ref_type == 'plq_id':
                    qty = select_print_qty(qty_ref, db_ref_type, db_ref)
                    label_text_handler(db_ref_type, db_ref, label_ref, printer, qty)
                elif db_ref_type == 'ord_no':
                    ord_no = db_ref
                    orl_ids = ord_no_orl_id(ord_no)
                    for orl_id in orl_ids:
                        qty = select_print_qty(qty_ref, 'orl_id', orl_id)
                        label_text_handler('orl_id', orl_id, label_ref, printer, qty)


main()
