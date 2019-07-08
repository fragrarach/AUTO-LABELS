import data
import statements
import printers
import files


def listen_task(config, notify):
    raw_payload = notify.payload

    db_ref, db_ref_type, label_type, label_name, qty_ref, station = data.payload_handler(raw_payload)

    printer = printers.select_printer(config, label_name, station)

    if db_ref_type == 'plq_id':
        plq_id = db_ref
        cli_no = statements.plq_id_cli_no(config, plq_id)

        customer = data.get_cli_no_customer(config, cli_no) if label_type != 'GENERIC' else 'generic'

        label_path = config.DYNAMIC_LABELS[customer][label_name]['path']
        qty = data.select_print_qty(config, qty_ref, db_ref_type, db_ref)
        files.dynamic_label_handler(config, db_ref_type, db_ref, label_path, customer, label_name, printer, qty)

    elif db_ref_type == 'ord_no':
        ord_no = db_ref
        cli_no = statements.ord_no_cli_no(config, ord_no)
        customer = data.get_cli_no_customer(config, cli_no)
        orl_ids = statements.ord_no_orl_id(config, ord_no)
        for orl_id in orl_ids:
            orl_id = orl_id[0]
            prt_no = statements.orl_id_prt_no(config, orl_id)
            qty = data.select_print_qty(config, qty_ref, 'orl_id', orl_id)
            if qty > 0:
                if not label_name:
                    label_name = f'{prt_no}'

                content_type = files.get_label_content_type(config, customer, label_name)
                if content_type == 'static':
                    label_path = config.LABEL_DIR + '\\' + customer + '\\' + content_type + '\\' + label_name + '.label'
                    printers.dymo_print(config, printer, label_path, qty)
                elif content_type == 'dynamic':
                    label_path = config.DYNAMIC_LABELS[customer][label_name]['path']
                    files.dynamic_label_handler(
                        config, 'orl_id', orl_id, label_path, customer, label_name, printer, qty)
