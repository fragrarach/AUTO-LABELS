import data
import statements
import printers
import files


def listen_task(config, notify):
    raw_payload = notify.payload

    db_ref, db_ref_type, label_ref, qty_ref, station = data.payload_handler(raw_payload)

    printer = printers.select_printer(label_ref, station)

    label = files.select_label(config, label_ref)

    if db_ref_type == 'plq_id':
        qty = data.select_print_qty(config, qty_ref, db_ref_type, db_ref)
        files.label_text_handler(config, db_ref_type, db_ref, label_ref, label, printer, qty)
    elif db_ref_type == 'ord_no':
        ord_no = db_ref
        orl_ids = statements.ord_no_orl_id(config, ord_no)
        for orl_id in orl_ids:
            qty = data.select_print_qty(config, qty_ref, 'orl_id', orl_id)
            files.label_text_handler(config, 'orl_id', orl_id, label_ref, label, printer, qty)