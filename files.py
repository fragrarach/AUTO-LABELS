import os
from shutil import copyfile
from statements import plq_id_prt_no, prt_no_prt_desc, orl_id_prt_no, orl_id_prt_desc, plq_id_plq_note
from data import serial_no_range
import printers


# Return label based on which report is being run
def select_label(config, label_ref, cli_no=None, orl_id=None):
    label_dir = config.PARENT_DIR + '\\files\\DYMO LABELS'

    if label_ref == 'CONTROL PANEL':
        return label_dir + r'\generic\dynamic\CONTROL PANEL.label'

    elif label_ref == 'UNIT':
        return label_dir + r'\generic\dynamic\UNIT.label'

    elif label_ref == 'SHIPPING':
        return label_dir + r'\generic\dynamic\SHIPPING.label'

    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return label_dir + r'\generic\dynamic\SHIPPING SERIAL NUMBER.label'

    elif label_ref == 'SERIAL NUMBER':
        return label_dir + r'\generic\dynamic\SERIAL NUMBER.label'

    elif label_ref == 'CLIENT':
        for customer in config.CUSTOMERS:
            if cli_no in config.CUSTOMERS[f'{customer}']:
                prt_no = orl_id_prt_no(config, orl_id)
                label_name = fr'\{prt_no}.label'
                static_label_path = label_dir + '\\' + customer + r'\static' + label_name
                dynamic_label_path = label_dir + '\\' + customer + r'\dynamic' + label_name

                if os.path.exists(static_label_path):
                    label_type = 'static'
                    return static_label_path, label_type

                elif os.path.exists(dynamic_label_path):
                    label_type = 'dynamic'
                    return dynamic_label_path, label_type
    return


# Create a working copy of label template
def copy_label_template(label):
    label_name = label.split('\\')[-1]
    label_dir = '\\'.join(label.split('\\')[:-1])
    label_template = f'{label_dir}\\TEMPLATES\\{label_name}'
    copyfile(label_template, label)


def control_panel_label_text(prt_no):
    pairs = []
    prt_no_pair = ['<Text></Text>', f'<Text>{prt_no}</Text>']
    pairs.append(prt_no_pair)

    return pairs


def unit_label_text(serial_no, prt_no):
    pairs = []
    serial_no = str(serial_no)
    serial_no_pair = ['<Text>1</Text>', f'<Text>{serial_no}</Text>']
    pairs.append(serial_no_pair)
    prt_no_pair = ['<Text>2</Text>', f'<Text>{prt_no}</Text>']
    pairs.append(prt_no_pair)

    return pairs


def serial_number_label_barcode_text(serial_no):
    pairs = []
    serial_no = str(serial_no)
    serial_no_pair = ['<Text></Text>', f'<Text>{serial_no}</Text>']
    pairs.append(serial_no_pair)

    return pairs


def serial_number_label_hybrid_text(serial_no, mod43):
    pairs = []
    serial_no = str(serial_no)
    serial_no_text_pair = ['<String xml:space="preserve">SERIAL_NUMBER</String>',
                           f'<String xml:space="preserve">{serial_no}</String>']
    pairs.append(serial_no_text_pair)
    serial_no_barcode_pair = ['{SERIAL_NUMBER}', f'{serial_no}{mod43}']
    pairs.append(serial_no_barcode_pair)

    return pairs


def shipping_serial_number_label_text(serial_no, prt_no, prt_desc):
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

    return pairs


def shipping_label_text(prt_no, prt_desc):
    pairs = []
    prt_no_barcode_pair = ['<Text>3</Text>', f'<Text>{prt_no}</Text>']
    pairs.append(prt_no_barcode_pair)
    prt_no_pair = ['<String xml:space="preserve">: 3</String>',
                   f'<String xml:space="preserve">: {prt_no}</String>']
    pairs.append(prt_no_pair)
    prt_desc_pair = ['<String xml:space="preserve">: 2</String>',
                     f'<String xml:space="preserve">: {prt_desc}</String>']
    pairs.append(prt_desc_pair)

    return pairs


# Edit label text based on label/DB reference passed by payload
def label_text_handler(config, db_ref_type, db_ref, label_ref, label, printer, qty):
    # PLETI Report
    if db_ref_type == 'plq_id':
        plq_id = db_ref
        prt_no = plq_id_prt_no(config, plq_id)
        prt_desc = prt_no_prt_desc(config, prt_no)
        plq_note = plq_id_plq_note(config, plq_id)
        serial_no_list = serial_no_range(plq_note)

        if label_ref == 'CONTROL PANEL':
            pairs = control_panel_label_text(prt_no)
            print(f'Printing {prt_no} on {label_ref} label using {printer}')
            printers.print_label(config, pairs, label, printer, qty)

        elif label_ref == 'UNIT':
            for serial_no in serial_no_list:
                pairs = unit_label_text(serial_no, prt_no)
                print(f'Printing {serial_no}, {prt_no} on {label_ref} label using {printer}')
                printers.print_label(config, pairs, label, printer)

        elif label_ref == 'SERIAL NUMBER':
            for serial_no in serial_no_list:
                pairs = serial_number_label_barcode_text(serial_no)
                print(f'Printing {serial_no} on {label_ref} label using {printer}')
                printers.print_label(config, pairs, label, printer, qty)

        elif label_ref == 'SHIPPING SERIAL NUMBER':
            for serial_no in serial_no_list:
                pairs = shipping_serial_number_label_text(serial_no, prt_no, prt_desc)
                print(f'Printing {prt_no}, {serial_no}, {prt_desc} on {label_ref} label using {printer}')
                printers.print_label(config, pairs, label, printer, qty)

    # CCETI Report
    elif db_ref_type == 'orl_id':
        orl_id = db_ref
        prt_no = orl_id_prt_no(config, orl_id)
        prt_desc = orl_id_prt_desc(config, orl_id)

        if label_ref == 'SHIPPING':
            pairs = shipping_label_text(prt_no, prt_desc)
            print(f'Printing {prt_no}, {prt_desc} on {label_ref} label using {printer}')
            printers.print_label(config, pairs, label, printer, qty)


# Pass list of pairs of strings (old text, new text) to edit label XML text
def insert_label_text(pairs, label):
    for pair in pairs:
        with open(label, "r") as file:
            file_data = file.read()
        file_data = file_data.replace(pair[0], pair[1])

        with open(label, "w") as file:
            file.write(file_data)
