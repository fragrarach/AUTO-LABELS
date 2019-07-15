import os
from shutil import copyfile
from statements import plq_id_prt_no, prt_no_prt_desc, orl_id_prt_no, orl_id_prt_desc, plq_id_plq_note, orl_id_plq_note
from data import serial_no_range, modulo_43
import printers


def get_label_content_type(config, payload):
    prefix = config.LABEL_DIR + '\\' + payload['customer']
    suffix = '\\' + payload['label_name'] + '.label'
    static_label_path = prefix + r'\static' + suffix
    dynamic_label_path = prefix + r'\dynamic' + suffix

    if os.path.exists(static_label_path):
        payload['label_path'] = static_label_path
        payload['content_type'] = 'static'

    elif os.path.exists(dynamic_label_path):
        payload['label_path'] = dynamic_label_path
        payload['content_type'] = 'dynamic'


# Create a working copy of label template
def copy_label_template(payload):
    label_name = payload['label_path'].split('\\')[-1]
    label_dir = '\\'.join(payload['label_path'].split('\\')[:-1])
    label_template = f"{label_dir}\\TEMPLATES\\{label_name}"
    copyfile(label_template, payload['label_path'])


def dynamic_label_handler(config, payload):
    # PLETI Report
    if payload['db_ref_type'] == 'plq_id':
        prt_no = plq_id_prt_no(config, payload['plq_id'])
        prt_desc = prt_no_prt_desc(config, prt_no)
        plq_note = plq_id_plq_note(config, payload['plq_id'])
        serial_no_list = serial_no_range(plq_note)

        if payload['label_name'] == 'CONTROL PANEL':
            payload['refs'] = [prt_no]
            printers.print_dynamic_label(config, payload)

        elif payload['label_name'] == 'UNIT':
            for serial_no in serial_no_list:
                payload['refs'] = [serial_no, prt_no]
                printers.print_dynamic_label(config, payload)

        elif payload['label_name'] == 'SERIAL NUMBER':
            for serial_no in serial_no_list:
                payload['refs'] = [serial_no]
                printers.print_dynamic_label(config, payload)

        elif payload['label_name'] == 'SHIPPING SERIAL NUMBER':
            for serial_no in serial_no_list:
                payload['refs'] = [prt_no, serial_no, prt_no, prt_desc, serial_no]
                printers.print_dynamic_label(config, payload)

    # CCETI Report
    elif payload['db_ref_type'] == 'orl_id':
        orl_id = payload['db_ref']
        prt_no = orl_id_prt_no(config, orl_id)
        prt_desc = orl_id_prt_desc(config, orl_id)

        if payload['customer'] == 'sirona':
            plq_note = orl_id_plq_note(config, orl_id)
            serial_no_list = serial_no_range(plq_note)
            for serial_no in serial_no_list:
                mod43 = modulo_43(serial_no)
                payload['refs'] = [serial_no, mod43, serial_no]
                payload['qty'] = 1
                printers.print_dynamic_label(config, payload)
        else:
            if payload['label_name'] == 'SHIPPING':
                payload['refs'] = [prt_no, prt_no, prt_desc]
                printers.print_dynamic_label(config, payload)


def insert_label_text(config, payload):
    for index, ref in enumerate(payload['refs']):
        with open(payload['label_path'], "r") as file:
            file_data = file.read()
        dynamic_text = config.DYNAMIC_LABELS[payload['customer']][payload['label_name']]['dynamic text'][index]
        file_data = file_data.replace(dynamic_text, str(ref), 1)

        with open(payload['label_path'], "w") as file:
            file.write(file_data)
