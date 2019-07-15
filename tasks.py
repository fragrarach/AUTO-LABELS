import data
import statements
import printers
import files


def listen_task(config, notify):
    raw_payload = notify.payload
    payload = data.payload_handler(raw_payload)
    payload['printer'] = printers.select_printer(config, payload)
    payload['customer'] = data.get_cli_no_customer(config, payload)

    if payload['db_ref_type'] == 'plq_id':
        payload['plq_id'] = payload['db_ref']
        payload['cli_no'] = statements.plq_id_cli_no(config, payload['plq_id'])
        payload['label_path'] = config.DYNAMIC_LABELS[payload['customer']][payload['label_name']]['path']
        payload['qty'] = data.select_print_qty(config, payload)

        files.dynamic_label_handler(config, payload)

    elif payload['db_ref_type'] == 'ord_no':
        ord_no = payload['db_ref']
        payload['cli_no'] = statements.ord_no_cli_no(config, ord_no)
        orl_ids = statements.ord_no_orl_id(config, ord_no)
        payload['db_ref_type'] = 'orl_id'
        for orl_id in orl_ids:
            payload['orl_id'] = orl_id[0]
            prt_no = statements.orl_id_prt_no(config, orl_id)
            payload['qty'] = data.select_print_qty(config, payload)
            if not payload['label_name']:
                payload['label_name'] = prt_no

            files.get_label_content_type(config, payload)
            if payload['content_type'] == 'static':
                printers.dymo_print(config, payload)
            elif payload['content_type'] == 'dynamic':
                files.dynamic_label_handler(config, payload)
