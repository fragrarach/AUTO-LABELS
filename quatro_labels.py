from printers import get_dymo_printers
from listen import listen


def main():
    get_dymo_printers()
    listen()


if __name__ == "__main__":
    main()
