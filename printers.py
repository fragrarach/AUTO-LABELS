from files import copy_label_template, insert_label_text
from quatro import log


# Pull list of currently online shared Dymo printers on the network
def get_dymo_printers(config):
    printers = config.DYMO_COM.GetDymoPrinters()
    log(printers)


# Return printer based on which computer is passing the payload
def select_printer(config, payload):
    if payload['station'] in ('SHIPPING03', 'CAD2', 'DESKTOP-FG8A5QJ', 'DESKTOP-69UD4SH'):
        shipping_printer = f"\\\\{config.STATIONS[payload['station']]['full computer name']}" \
                           f"\\{config.STATIONS[payload['station']]['shipping printer name']}"
    else:
        shipping_printer = config.PRINTERS['default shipping printer']

    if payload['label_name'] == 'CONTROL PANEL':
        return config.PRINTERS['control panel']

    elif payload['label_name'] == 'UNIT':
        return config.PRINTERS['unit printer']

    elif payload['label_name'] in ('SHIPPING', 'SHIPPING SERIAL NUMBER', 'CLIENT'):
        return shipping_printer

    elif payload['label_name'] == 'SERIAL NUMBER':
        return config.PRINTERS['serial number printer']


# Dymo COM API functions
def dymo_print(config, payload):
    log(f"Printing {payload['qty']} {payload['label_path']} on {payload['printer']}")
    config.DYMO_COM.SelectPrinter(payload['printer'])
    config.DYMO_COM.Open(payload['label_path'])
    config.DYMO_COM.StartPrintJob()
    config.DYMO_COM.Print(payload['qty'], False)
    config.DYMO_COM.EndPrintJob()
    log(f'Print job completed')


def print_dynamic_label(config, payload):
    copy_label_template(payload)
    insert_label_text(config, payload)
    dymo_print(config, payload)
