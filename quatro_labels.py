from printers import get_dymo_printers
from quatro import listen, init_app_log_dir, log
import config
from tasks import listen_task


def main():
    init_app_log_dir()
    log(f'Starting {__file__}')
    label_config = config.Config()
    listen(label_config, listen_task, get_dymo_printers(label_config))


if __name__ == "__main__":
    main()
