from win32com.client import Dispatch
from os.path import dirname, abspath
from quatro import sigm_connect


class Config:
    LISTEN_CHANNEL = 'labels'
    
    # Dymo COM configs
    DYMO_COM = Dispatch('Dymo.DymoAddIn')
    DYMO_LABEL = Dispatch('Dymo.DymoLabels')

    PARENT_DIR = dirname(abspath(__file__))

    STATIONS = {
        'CAD2': {
            'full computer name': 'CAD2.aerofil.local',
            'shipping printer name': 'DEV SHIPPING DYMO'
        },
        'SHIPPING03': {
            'full computer name': 'SHIPPING03',
            'shipping printer name': 'ERICK DYMO',
            'serial printer name': 'DYMO 450 SERIAL'
        },
        'DESKTOP-FG8A5QJ': {
            'full computer name': 'DESKTOP-FG8A5QJ.aerofil.local',
            'shipping printer name': 'GATTA DYMO'
        },
        'DESKTOP-69UD4SH': {
            'full computer name': 'DESKTOP-69UD4SH',
            'shipping printer name': 'JORGE DYMO'
        }
    }

    PRINTERS = {
        'control panel': 'DEV DYMO',
        'unit printer': 'DEV DYMO',
        'default shipping printer': f"\\\\{STATIONS['DESKTOP-FG8A5QJ']['full computer name']}"
                                    f"\\{STATIONS['DESKTOP-FG8A5QJ']['shipping printer name']}",
        'serial number printer': f"\\\\{STATIONS['SHIPPING03']['full computer name']}"
                                 f"\\{STATIONS['SHIPPING03']['serial printer name']}"
    }

    CUSTOMERS = {
        'sirona': [
            'RS109',
            'RS109U',
            'RES109',
            'TEST_CLI'
        ]
    }

    LABELS = {
        'control panel': [
            'prt_no'
        ],
        'unit': [
            'serial_no',
            'prt_no'
        ],
        'serial number': [
            'serial_no'
        ],
        'shipping serial number': [
            'prt_no',
            'serial_no',
            'prt_no',
            'prt_desc',
            'serial_no'
        ],
        'shipping': [
            'prt_no',
            'prt_no',
            'prt_desc'
        ]
    }

    def __init__(self):
        self.sigm_connection, self.sigm_db_cursor = sigm_connect(Config.LISTEN_CHANNEL)
