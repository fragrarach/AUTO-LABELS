from printers import get_dymo_printers
from quatro import listen, init_app_log_dir, log, configuration as c
from config import Config
from tasks import listen_task


def main():
    c.config = Config(__file__)
    init_app_log_dir()
    log(f'Starting {__file__}')
    c.config.sql_connections()
    listen(listen_task, else_task=get_dymo_printers())


if __name__ == "__main__":
    main()
