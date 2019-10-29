from files import copy_label_template, insert_label_text
from quatro import log, configuration as c


# Pull list of currently online shared Dymo printers on the network
def get_dymo_printers():
    printers = c.config.DYMO_COM.GetDymoPrinters()
    log(printers)


# Return printer based on which computer is passing the payload
def select_printer(payload):
    if payload['station'] in ('SHIPPING03', 'CAD2', 'DESKTOP-FG8A5QJ', 'DESKTOP-69UD4SH'):
        shipping_printer = f"\\\\{c.config.STATIONS[payload['station']]['full computer name']}" \
                           f"\\{c.config.STATIONS[payload['station']]['shipping printer name']}"
    else:
        shipping_printer = c.config.PRINTERS['default shipping printer']

    if payload['label_name'] == 'CONTROL PANEL':
        return c.config.PRINTERS['control panel']

    elif payload['label_name'] == 'UNIT':
        return c.config.PRINTERS['unit printer']

    elif payload['label_name'] in ('SHIPPING', 'SHIPPING SERIAL NUMBER', 'CLIENT'):
        return shipping_printer

    elif payload['label_name'] == 'SERIAL NUMBER':
        return c.config.PRINTERS['serial number printer']


# Dymo COM API functions
def dymo_print(payload):
    log(f"Printing {payload['qty']} {payload['label_path']} on {payload['printer']}")
    c.config.DYMO_COM.SelectPrinter(payload['printer'])
    c.config.DYMO_COM.Open(payload['label_path'])
    c.config.DYMO_COM.StartPrintJob()
    c.config.DYMO_COM.Print(payload['qty'], False)
    c.config.DYMO_COM.EndPrintJob()
    log(f'Print job completed')


def print_dynamic_label(payload):
    copy_label_template(payload)
    insert_label_text(payload)
    dymo_print(payload)
