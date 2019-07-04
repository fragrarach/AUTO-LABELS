import os
import data
import statements
import printers
import files


def listen_task(config, notify):
    raw_payload = notify.payload

    db_ref, db_ref_type, label_ref, qty_ref, station = data.payload_handler(raw_payload)

    printer = printers.select_printer(config, label_ref, station)

    if db_ref_type == 'plq_id':
        label = files.select_label(config, label_ref, db_ref)
        qty = data.select_print_qty(config, qty_ref, db_ref_type, db_ref)
        files.label_text_handler(config, db_ref_type, db_ref, label_ref, label, printer, qty)

    elif db_ref_type == 'ord_no':
        ord_no = db_ref
        cli_no = statements.ord_no_cli_no(config, ord_no)
        orl_ids = statements.ord_no_orl_id(config, ord_no)
        for orl_id in orl_ids:
            orl_id = orl_id[0]
            if files.select_label(config, label_ref, cli_no, orl_id):
                label, label_type = files.select_label(config, label_ref, cli_no, orl_id)
                qty = data.select_print_qty(config, qty_ref, 'orl_id', orl_id)

                if qty > 0:
                    if label_ref == 'CLIENT' and os.path.exists(label):
                        if label_type == 'static':
                            print(f'Printing {label} on {printer} from {station}')
                            printers.dymo_print(config, printer, label, qty)
                        # Assuming all dynamic labels are serial number only, add sub types if this changes
                        elif label_type == 'dynamic':
                            plq_note = statements.orl_id_plq_note(config, orl_id)
                            serial_no_list = data.serial_no_range(plq_note)
                            for serial_no in serial_no_list:
                                mod43 = data.modulo_43(serial_no)
                                pairs = files.serial_number_label_hybrid_text(serial_no, mod43)
                                qty = 1
                                printers.print_label(config, pairs, label, printer, qty)

                    elif label_ref == 'SHIPPING':
                        files.label_text_handler(config, 'orl_id', orl_id, label_ref, label, printer, qty)
