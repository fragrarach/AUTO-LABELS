import re
from sigm import sigm_connect
from config import Config
from printers import get_dymo_printers, select_printer
from data import select_print_qty
from files import label_text_handler, select_label
from sql import ord_no_orl_id


# Split payload string, return named variables
def payload_handler(payload):
    db_ref = payload.split(', ')[0]
    db_ref_type = payload.split(', ')[1]
    label_ref = payload.split(', ')[2]
    qty_ref = payload.split(', ')[3]
    sigm_string = payload.split(', ')[4]
    station = re.findall(r'(?<= w)(.*)$', sigm_string)[0]

    return db_ref, db_ref_type, label_ref, qty_ref, station


def listen():
    while 1:
        try:
            Config.SIGM_CONNECTION.poll()
        except:
            print('Database cannot be accessed, PostgreSQL service probably rebooting')
            try:
                Config.SIGM_CONNECTION.close()
                Config.SIGM_CONNECTION, Config.SIGM_DB_CURSOR = sigm_connect(Config.LISTEN_CHANNEL)

            except:
                pass
            else:
                get_dymo_printers()
        else:
            Config.SIGM_CONNECTION.commit()
            while Config.SIGM_CONNECTION.notifies:
                notify = Config.SIGM_CONNECTION.notifies.pop()
                raw_payload = notify.payload

                db_ref, db_ref_type, label_ref, qty_ref, station = payload_handler(raw_payload)

                printer = select_printer(label_ref, station)

                label = select_label(label_ref)

                if db_ref_type == 'plq_id':
                    qty = select_print_qty(qty_ref, db_ref_type, db_ref)
                    label_text_handler(db_ref_type, db_ref, label_ref, label, printer, qty)
                elif db_ref_type == 'ord_no':
                    ord_no = db_ref
                    orl_ids = ord_no_orl_id(ord_no)
                    for orl_id in orl_ids:
                        qty = select_print_qty(qty_ref, 'orl_id', orl_id)
                        label_text_handler('orl_id', orl_id, label_ref, label, printer, qty)


if __name__ == "__main__":
    listen()
