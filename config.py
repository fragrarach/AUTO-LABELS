from win32com.client import Dispatch
from os.path import dirname, abspath
from quatro import sigm_connect


class Config:
    def __init__(self, main_file_path):
        self.main_file_path = main_file_path
        self.parent_dir = dirname(abspath(main_file_path))
        self.sigm_connection = None
        self.sigm_db_cursor = None
        self.LISTEN_CHANNEL = 'labels'

    def sql_connections(self):
        self.sigm_connection, self.sigm_db_cursor = sigm_connect(self.LISTEN_CHANNEL)

    # Dymo COM configs
    DYMO_COM = Dispatch('Dymo.DymoAddIn')
    DYMO_LABEL = Dispatch('Dymo.DymoLabels')

    PARENT_DIR = dirname(abspath(__file__))
    LABEL_DIR = PARENT_DIR + r'\files\DYMO LABELS'

    STATIONS = {
        'CAD2': {
            'full computer name': 'CAD2',
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

    DYNAMIC_LABELS = {
        'generic': {
            'CONTROL PANEL': {
                'dynamic text': [
                    'prt_no'
                ],
                'path': LABEL_DIR + r'\generic\dynamic\CONTROL_PANEL.label'
            },
            'UNIT': {
                'dynamic text': [
                    'serial_no',
                    'prt_no'
                ],
                'path': LABEL_DIR + r'\generic\dynamic\UNIT.label'
            },
            'SERIAL NUMBER': {
                'dynamic text': [
                    'serial_no'
                ],
                'path': LABEL_DIR + r'\generic\dynamic\SERIAL_NUMBER.label'
            },
            'SHIPPING SERIAL NUMBER': {
                'dynamic text': [
                    'prt_no',
                    'serial_no',
                    'prt_no',
                    'prt_desc',
                    'serial_no'
                ],
                'path': LABEL_DIR + r'\generic\dynamic\SHIPPING_SERIAL_NUMBER.label'
            },
            'SHIPPING': {
                'dynamic text': [
                    'prt_no',
                    'prt_no',
                    'prt_desc'
                ],
                'path': LABEL_DIR + r'\generic\dynamic\SHIPPING.label'
            }
        },
        'sirona': {
            'AX139-16 R4.3': {
                'dynamic text': [
                    'serial_no',
                    'mod_43',
                    'serial_no'
                ],
                'path': LABEL_DIR + r'\sirona\dynamic\AX139-16 R4.3.label'
            },
            'AX139-25 R4.3': {
                'dynamic text': [
                    'serial_no',
                    'mod_43',
                    'serial_no'
                ],
                'path': LABEL_DIR + r'\sirona\dynamic\AX139-25 R4.3.label'
            }
        }
    }
