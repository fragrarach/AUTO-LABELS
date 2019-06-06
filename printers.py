from files import copy_label_template, insert_label_text


# Pull list of currently online shared Dymo printers on the network
def get_dymo_printers(config):
    printers = config.DYMO_COM.GetDymoPrinters()
    print(printers)


# Return printer based on which computer is passing the payload
def select_printer(label_ref, station):
    control_panel_printer = r'DEV DYMO'
    unit_printer = r'DEV DYMO'
    serial_number_printer = f'\\\\SHIPPING03\\DYMO 450 SERIAL'

    if station == 'SHIPPING03':
        shipping_printer = f'\\\\SHIPPING03\\ERICK DYMO'
    elif station == 'CAD2':
        shipping_printer = f'\\\\CAD2.aerofil.local\\DEV SHIPPING DYMO'
    elif station == 'DESKTOP-FG8A5QJ':
        shipping_printer = f'\\\\DESKTOP-FG8A5QJ.aerofil.local\\GATTA DYMO'
    elif station == 'DESKTOP-69UD4SH':
        shipping_printer = f'\\\\DESKTOP-69UD4SH\\JORGE DYMO'
    else:
        shipping_printer = f'\\\\DESKTOP-FG8A5QJ.aerofil.local\\GATTA DYMO'

    if label_ref == 'CONTROL PANEL':
        return control_panel_printer
    elif label_ref == 'UNIT':
        return unit_printer
    elif label_ref == 'SHIPPING':
        return shipping_printer
    elif label_ref == 'SHIPPING SERIAL NUMBER':
        return shipping_printer
    elif label_ref == 'SERIAL NUMBER':
        return serial_number_printer


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
