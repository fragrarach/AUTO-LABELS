import data
import statements
import printers
import files
from quatro import log, configuration as c


def listen_task(notify):
    try:
        raw_payload = notify.payload
        payload = data.payload_handler(raw_payload)
        payload['printer'] = printers.select_printer(payload)

        if payload['db_ref_type'] == 'plq_id':
            payload['plq_id'] = payload['db_ref']
            payload['cli_no'] = statements.plq_id_cli_no(payload['plq_id'])
            payload['customer'] = data.get_cli_no_customer(payload)
            payload['label_path'] = c.config.DYNAMIC_LABELS[payload['customer']][payload['label_name']]['path']
            payload['qty'] = data.select_print_qty(payload)

            files.dynamic_label_handler(payload)

        elif payload['db_ref_type'] == 'ord_no':
            ord_no = payload['db_ref']
            payload['cli_no'] = statements.ord_no_cli_no(ord_no)
            payload['customer'] = data.get_cli_no_customer(payload)
            orl_ids = statements.ord_no_orl_id(ord_no)
            payload['db_ref_type'] = 'orl_id'
            for orl_id in orl_ids:
                payload['orl_id'] = orl_id[0]
                prt_no = statements.orl_id_prt_no(payload['orl_id'])
                payload['qty'] = data.select_print_qty(payload)
                if not payload['label_name']:
                    payload['label_name'] = prt_no

                files.get_label_content_type(payload)
                if payload['content_type'] == 'static':
                    printers.dymo_print(payload)
                elif payload['content_type'] == 'dynamic':
                    files.dynamic_label_handler(payload)
    except Exception as e:
        log(getattr(e, 'message', repr(e)))
