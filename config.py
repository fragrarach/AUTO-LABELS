from win32com.client import Dispatch
from os.path import dirname, abspath
from quatro import sigm_connect


class Config:
    LISTEN_CHANNEL = 'labels'
    
    # Dymo COM configs
    DYMO_COM = Dispatch('Dymo.DymoAddIn')
    DYMO_LABEL = Dispatch('Dymo.DymoLabels')

    PARENT_DIR = dirname(abspath(__file__))

    def __init__(self):
        self.sigm_connection, self.sigm_db_cursor = sigm_connect(Config.LISTEN_CHANNEL)
