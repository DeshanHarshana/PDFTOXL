"""
Entry point for the PDF → Excel converter application.

Usage
-----
    python main.py
"""

from gui import App


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
