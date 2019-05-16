from printers import get_dymo_printers
from quatro import listen
import config
from tasks import listen_task


def main():
    label_config = config.Config()
    get_dymo_printers(label_config)
    listen(label_config, listen_task, get_dymo_printers(label_config))


if __name__ == "__main__":
    main()
