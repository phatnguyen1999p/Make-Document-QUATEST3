"""
QUATEST3 - HỆ THỐNG XỬ LÝ VĂN BẢN (Word, Excel, PDF)
Main application entry point.
"""

import sys
import tkinter as tk
from gui.app import UnifiedDocumentSystem


def main():
    root = tk.Tk()
    app = UnifiedDocumentSystem(root)
    root.mainloop()


if __name__ == "__main__":
    main()
