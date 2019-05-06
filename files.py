from shutil import copyfile
from config import Config
from sql import plq_id_prt_no, prt_no_prt_desc, orl_id_prt_no, orl_id_prt_desc
from data import serial_no_range
import printers


# Return label based on which report is being run
def select_label(label_ref):
    label_dir = Config.PARENT_DIR + '\\files\\DYMO LABELS'

    if label_ref == 'CONTROL PANEL':
        return label_dir + r'\CONTROL PANEL.label'
    elif label_ref == 'UNIT':
        return label_dir + r'\UNIT.label'
    elif label_ref == 'SHIPPING':
        return label_dir + r'\SHIPPING.label'
    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return label_dir + r'\SHIPPING SERIAL NUMBER.label'
    elif label_ref == 'SERIAL NUMBER':
        return label_dir + r'\SERIAL NUMBER.label'


# Create a working copy of label template
def copy_label_template(label):
    label_name = label.split('\\')[-1]
    label_dir = '\\'.join(label.split('\\')[:-1])
    label_template = f'{label_dir}\\TEMPLATES\\{label_name}'
    copyfile(label_template, label)


# Edit label text based on label/DB reference passed by payload
def label_text_handler(db_ref_type, db_ref, label_ref, label, printer, qty):
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

            print(f'Printing {prt_no} on {label_ref} label using {printer}')
            printers.print_label(pairs, label, printer, qty)

        elif label_ref == 'UNIT':
            for serial_no in serial_no_list:
                pairs = []
                serial_no = str(serial_no)
                serial_no_pair = ['<Text>1</Text>', f'<Text>{serial_no}</Text>']
                pairs.append(serial_no_pair)
                prt_no_pair = ['<Text>2</Text>', f'<Text>{prt_no}</Text>']
                pairs.append(prt_no_pair)

                print(f'Printing {serial_no}, {prt_no} on {label_ref} label using {printer}')
                printers.print_label(pairs, label, printer)

        elif label_ref == 'SERIAL NUMBER':
            for serial_no in serial_no_list:
                pairs = []
                serial_no = str(serial_no)
                serial_no_pair = ['<Text></Text>', f'<Text>{serial_no}</Text>']
                pairs.append(serial_no_pair)

                print(f'Printing {serial_no} on {label_ref} label using {printer}')
                printers.print_label(pairs, label, printer, qty)

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

                print(f'Printing {prt_no}, {serial_no}, {prt_desc} on {label_ref} label using {printer}')
                printers.print_label(pairs, label, printer, qty)

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

            print(f'Printing {prt_no}, {prt_desc} on {label_ref} label using {printer}')
            printers.print_label(pairs, label, printer, qty)


# Pass list of pairs of strings (old text, new text) to edit label XML text
def insert_label_text(pairs, label):
    for pair in pairs:
        with open(label, "r") as file:
            file_data = file.read()
        file_data = file_data.replace(pair[0], pair[1])

        with open(label, "w") as file:
            file.write(file_data)
