import psycopg2.extensions
from win32com.client import Dispatch
from os.path import dirname, abspath
from sigm import sigm_connect, log_connect


# PostgreSQL DB connection configs
psycopg2.extensions.register_type(psycopg2.extensions.UNICODE)
psycopg2.extensions.register_type(psycopg2.extensions.UNICODEARRAY)


class Config:
    LISTEN_CHANNEL = 'labels'

    SIGM_CONNECTION, SIGM_DB_CURSOR = sigm_connect(LISTEN_CHANNEL)

    # Dymo COM configs
    DYMO_COM = Dispatch('Dymo.DymoAddIn')
    DYMO_LABEL = Dispatch('Dymo.DymoLabels')

    PARENT_DIR = dirname(abspath(__file__))
