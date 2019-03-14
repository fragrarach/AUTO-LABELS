import psycopg2.extensions
import re
import datetime
import os
import time
from win32com.client import Dispatch
from os.path import dirname, abspath

psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)

conn_sigm = psycopg2.connect("host='192.168.0.250' dbname='QuatroAir' user='SIGM' port='5493'")
conn_sigm.set_client_encoding('latin1')
conn_sigm.set_isolation_level(psycopg2.extensions.ISOLATION_LEVEL_AUTOCOMMIT)

sigm_listen = conn_sigm.cursor()
sigm_listen.execute('LISTEN labels;')
sigm_query = conn_sigm.cursor()

labelCom = Dispatch('Dymo.DymoAddIn')
labelText = Dispatch('Dymo.DymoLabels')


def plq_id_prt_no(plq_id):
    sql_exp = f'SELECT trim(prt_no) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            prt_no = cell
            return prt_no


def plq_id_plq_qty_per(plq_id):
    sql_exp = f'SELECT plq_qty_per FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            plq_qty_per = cell
            return plq_qty_per


def plq_id_plq_note(plq_id):
    sql_exp = f'SELECT trim(plq_note) FROM planning_lot_quantity WHERE plq_id = {plq_id}'
    sigm_query.execute(sql_exp)
    result_set = sigm_query.fetchall()

    for row in result_set:
        for cell in row:
            plq_note = cell
            return plq_note


def serial_no_range(plq_id):
    plq_note = plq_id_plq_note(plq_id)
    serial_numbers = plq_note.strip('SN: ')

    sn_list = []

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


def select_qty(qty_ref, db_ref_type, ref):
    if qty_ref == 'SINGLE':
        qty = 1
        return qty
    elif qty_ref == 'MULTI':
        if db_ref_type == 'plq_id':
            plq_id = ref
            qty = plq_id_plq_qty_per(plq_id)
            return qty


def select_printer(label_ref, station):
    control_panel_printer = r'DEV DYMO'
    unit_printer = r'DEV DYMO'
    shipping_printer = r'DEV DYMO'
    # serial_number_printer = f'\\\\{station}.aerofil.local\\SERIAL DYMO'
    serial_number_printer = f'\\\\DESKTOP-PKO4ES1.aerofil.local\\DYMO 450 SERIAL'

    if label_ref == 'CONTROL PANEL':
        return control_panel_printer
    elif label_ref == 'UNIT':
        return unit_printer
    elif label_ref == 'SHIPPING':
        return shipping_printer
    elif label_ref == 'SERIAL NUMBER':
        return serial_number_printer


def select_label(label_ref):
    label_dir = \
        r'E:\DATA\Fortune\SIGMWIN.DTA\QuatroAir\Documents\REFERENCE FILES\AUTOMATED LABEL GENERATION\DYMO LABELS'

    control_panel_label = label_dir + '\CONTROL PANEL.label'
    unit_label = label_dir + r'\UNIT.label'
    shipping_label = label_dir + r'\SHIPPING.label'
    serial_number_label = label_dir + r'\SERIAL NUMBER.label'

    if label_ref == 'CONTROL PANEL':
        return control_panel_label
    elif label_ref == 'UNIT':
        return unit_label
    elif label_ref == 'SHIPPING':
        return shipping_label
    elif label_ref == 'SERIAL NUMBER':
        return serial_number_label


def print_label(db_ref_type, db_ref, label_ref, printer, qty):
    label = select_label(label_ref)
    if db_ref_type == 'plq_id':
        # PLETI Report
        plq_id = db_ref
        prt_no = plq_id_prt_no(plq_id)
        sn_list = serial_no_range(plq_id)

        if label_ref == 'CONTROL PANEL':
            insert_cp_label_text(prt_no, label)
            dymo_print(printer, label, qty)
            reset_cp_label_text(prt_no, label)

        elif label_ref == 'UNIT':
            for sn in sn_list:
                insert_unit_label_text(sn, prt_no, label)
                dymo_print(printer, label)
                reset_unit_label_text(sn, prt_no, label)

        elif label_ref == 'SERIAL NUMBER':
            for sn in sn_list:
                sn = str(sn)
                insert_cp_label_text(sn, label)
                dymo_print(printer, label, qty)
                reset_cp_label_text(sn, label)

    elif db_ref_type == 'ord_no':
        # CCETI Report
        print()


def dymo_print(printer, label, print_qty=1):
    labelCom.SelectPrinter(printer)
    labelCom.Open(label)
    labelCom.StartPrintJob()
    labelCom.Print(print_qty, False)
    labelCom.EndPrintJob()


def insert_cp_label_text(text, label):
    with open(label, "r") as file:
        file_data = file.read()
    file_data = file_data.replace('<Text></Text>', f'<Text>{text}</Text>')

    with open(label, "w") as file:
        file.write(file_data)


def reset_cp_label_text(text, label):
    with open(label, "r") as file:
        file_data = file.read()
    file_data = file_data.replace(f'<Text>{text}</Text>', '<Text></Text>')

    with open(label, "w") as file:
        file.write(file_data)


def insert_unit_label_text(sn_text, pn_text, label):
    with open(label, "r") as file:
        file_data = file.read()
    file_data = file_data.replace('<Text>1</Text>', f'<Text>{sn_text}</Text>')
    file_data = file_data.replace('<Text>2</Text>', f'<Text>{pn_text}</Text>')

    with open(label, "w") as file:
        file.write(file_data)


def reset_unit_label_text(sn_text, pn_text, label):
    with open(label, "r") as file:
        file_data = file.read()
    file_data = file_data.replace(f'<Text>{sn_text}</Text>', '<Text>1</Text>')
    file_data = file_data.replace(f'<Text>{pn_text}</Text>', '<Text>2</Text>')

    with open(label, "w") as file:
        file.write(file_data)


def main():
    while 1:
        conn_sigm.poll()
        conn_sigm.commit()
        while conn_sigm.notifies:
            notify = conn_sigm.notifies.pop()
            raw_payload = notify.payload
            print(raw_payload)
            db_ref = raw_payload.split(', ')[0]
            db_ref_type = raw_payload.split(', ')[1]
            label_ref = raw_payload.split(', ')[2]
            qty_ref = raw_payload.split(', ')[3]
            sigm_string = raw_payload.split(', ')[4]
            station = re.findall(r'(?<= w)(.*)$', sigm_string)[0]

            printer = select_printer(label_ref, station)

            if db_ref_type == 'plq_id':
                qty = select_qty(qty_ref, db_ref_type, db_ref)
                print_label(db_ref_type, db_ref, label_ref, printer, qty)
            elif db_ref_type == 'ord_no':
                pass


main()
