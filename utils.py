# utils.py
import tkinter as tk
from tkinter import ttk
import os
import sys
import platform
from pathlib import Path


def get_app_data_dir(app_name: str = 'Fruzy') -> str:
    """Return a platform-appropriate directory for persistent application data.

    macOS: ~/Library/Application Support/<app_name>
    Windows: %APPDATA%\\<app_name>
    Linux: ~/.local/share/<app_name>

    Ensures the directory exists and returns its string path.
    """
    system = platform.system()
    if system == 'Darwin':
        base = Path.home() / 'Library' / 'Application Support'
    elif system == 'Windows':
        base = Path(os.getenv('APPDATA', Path.home() / 'AppData' / 'Roaming'))
    else:
        base = Path.home() / '.local' / 'share'

    app_dir = base / app_name
    try:
        app_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Fallback to current directory
        return os.path.abspath('.')
    return str(app_dir)


def make_treeview(parent, columns, headings, widths=None, height=10):
    """Create and return a configured Treeview with scrollbar and extended selection."""
    frame = tk.Frame(parent)
    frame.pack(fill='both', expand=True)
    
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side='right', fill='y')
    
    tree = ttk.Treeview(
        frame,
        columns=columns,
        show='headings',
        height=height,
        yscrollcommand=scrollbar.set,
        selectmode='extended'  # ← Enables Shift+Click, range selection
    )
    tree.pack(side='left', fill='both', expand=True)
    scrollbar.config(command=tree.yview)
    
    for col, header in zip(columns, headings):
        tree.heading(col, text=header)
        width = widths[columns.index(col)] if widths else 120
        tree.column(col, width=width)
    
    return tree


def enable_treeview_select_all(tree_widget):
    """
    Enable Cmd+A (macOS) and Ctrl+A (Windows/Linux) to select all items
    in a ttk.Treeview with selectmode='extended'.
    
    Usage in any tab file:
        from utils import make_treeview, enable_treeview_select_all
        tree = make_treeview(...)
        enable_treeview_select_all(tree)
    """
    def _select_all(event=None):
        children = tree_widget.get_children()
        if children:
            tree_widget.selection_set(children)
        # Prevent default behavior (e.g., text cursor movement)
        return "break"
    
    tree_widget.bind('<Command-a>', _select_all)
    tree_widget.bind('<Control-a>', _select_all)