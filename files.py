import os
from shutil import copyfile
from statements import plq_id_prt_no, prt_no_prt_desc, orl_id_prt_no, orl_id_prt_desc, plq_id_plq_note, orl_id_plq_note
from data import serial_no_range, modulo_43
import printers


def get_label_content_type(config, customer, label_name):
    static_label_path = config.LABEL_DIR + '\\' + customer + r'\static' + '\\' + label_name + '.label'
    dynamic_label_path = config.LABEL_DIR + '\\' + customer + r'\dynamic' + '\\' + label_name + '.label'

    if os.path.exists(static_label_path):
        content_type = 'static'
        return content_type

    elif os.path.exists(dynamic_label_path):
        content_type = 'dynamic'
        return content_type
    return


# Create a working copy of label template
def copy_label_template(label):
    label_name = label.split('\\')[-1]
    label_dir = '\\'.join(label.split('\\')[:-1])
    label_template = f'{label_dir}\\TEMPLATES\\{label_name}'
    copyfile(label_template, label)


def dynamic_label_handler(config, db_ref_type, db_ref, label_path, customer, label_name, printer, qty):
    # PLETI Report
    if db_ref_type == 'plq_id':
        plq_id = db_ref
        prt_no = plq_id_prt_no(config, plq_id)
        prt_desc = prt_no_prt_desc(config, prt_no)
        plq_note = plq_id_plq_note(config, plq_id)
        serial_no_list = serial_no_range(plq_note)

        if label_name == 'CONTROL PANEL':
            refs = [prt_no]
            printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)

        elif label_name == 'UNIT':
            for serial_no in serial_no_list:
                refs = [serial_no, prt_no]
                printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)

        elif label_name == 'SERIAL NUMBER':
            for serial_no in serial_no_list:
                refs = [serial_no]
                printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)

        elif label_name == 'SHIPPING SERIAL NUMBER':
            for serial_no in serial_no_list:
                refs = [prt_no, serial_no, prt_no, prt_desc, serial_no]
                printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)

    # CCETI Report
    elif db_ref_type == 'orl_id':
        orl_id = db_ref
        prt_no = orl_id_prt_no(config, orl_id)
        prt_desc = orl_id_prt_desc(config, orl_id)

        if customer == 'sirona':
            plq_note = orl_id_plq_note(config, orl_id)
            serial_no_list = serial_no_range(plq_note)
            for serial_no in serial_no_list:
                mod43 = modulo_43(serial_no)
                refs = [serial_no, mod43, serial_no]
                qty = 1
                printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)
        else:
            if label_name == 'SHIPPING':
                refs = [prt_no, prt_no, prt_desc]
                printers.print_dynamic_label(config, refs, label_path, customer, label_name, printer, qty)


def insert_label_text(config, refs, label_path, customer, label_name):
    for index, ref in enumerate(refs):
        with open(label_path, "r") as file:
            file_data = file.read()
        dynamic_text = config.DYNAMIC_LABELS[customer][label_name]['dynamic text'][index]
        file_data = file_data.replace(dynamic_text, str(ref), 1)

        with open(label_path, "w") as file:
            file.write(file_data)
