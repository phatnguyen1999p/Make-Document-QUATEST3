import sys
from cli import run_cli
from gui import launch_gui


def main():
    # If no arguments provided or explicitly requesting GUI
    if len(sys.argv) == 1 or "--gui" in sys.argv:
        launch_gui()
    else:
        sys.exit(run_cli())


if __name__ == "__main__":
    main()
