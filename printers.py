from files import copy_label_template, insert_label_text


# Pull list of currently online shared Dymo printers on the network
def get_dymo_printers(config):
    printers = config.DYMO_COM.GetDymoPrinters()
    print(printers)


# Return printer based on which computer is passing the payload
def select_printer(config, label_ref, station):
    if station in ('SHIPPING03', 'CAD2', 'DESKTOP-FG8A5QJ', 'DESKTOP-69UD4SH'):
        shipping_printer = f"\\\\{config.STATIONS[station]['full computer name']}" \
                           f"\\{config.STATIONS[station]['shipping printer name']}"
    else:
        shipping_printer = config.PRINTERS['default shipping printer']

    if label_ref == 'CONTROL PANEL':
        return config.PRINTERS['control panel']
    elif label_ref == 'UNIT':
        return config.PRINTERS['unit printer']
    elif label_ref == 'SHIPPING':
        return shipping_printer
    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return shipping_printer
    elif label_ref == 'SERIAL NUMBER':
        return config.PRINTERS['serial number printer']


# Dymo COM API functions
def dymo_print(config, printer, label, print_qty=1):
    config.DYMO_COM.SelectPrinter(printer)
    config.DYMO_COM.Open(label)
    config.DYMO_COM.StartPrintJob()
    config.DYMO_COM.Print(print_qty, False)
    config.DYMO_COM.EndPrintJob()


# Set text label text, print label, revert label text
def print_label(config, pairs, label, printer, qty=1):
    copy_label_template(label)
    insert_label_text(pairs, label)
    dymo_print(config, printer, label, qty)
