"""
Progress dialog with animated spinner for long-running operations.
"""

import tkinter as tk
from tkinter import ttk
import threading
import time


class ProgressDialog:
    """A modal dialog with animated spinner for showing progress."""
    
    def __init__(self, parent, title="Processing..."):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("350x120")
        self.dialog.resizable(False, False)
        
        # Make it modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.dialog.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        # Prevent closing with X button
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # Create UI
        frame = ttk.Frame(self.dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Spinner animation
        self.spinner_label = ttk.Label(frame, text="", font=('Segoe UI', 14))
        self.spinner_label.pack(pady=(10, 5))
        
        # Status text
        self.status_label = ttk.Label(frame, text="Processing, please wait...", font=('Segoe UI', 10))
        self.status_label.pack()
        
        # Animation state
        self.running = True
        self.spinner_chars = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        self.spinner_index = 0
        
        # Start animation
        self._animate()
    
    def _animate(self):
        """Animate the spinner."""
        if self.running:
            self.spinner_label.config(text=self.spinner_chars[self.spinner_index])
            self.spinner_index = (self.spinner_index + 1) % len(self.spinner_chars)
            self.dialog.after(100, self._animate)
    
    def update_status(self, text):
        """Update the status text."""
        self.status_label.config(text=text)
        self.dialog.update_idletasks()
    
    def close(self):
        """Close the dialog."""
        self.running = False
        self.dialog.grab_release()
        self.dialog.destroy()


def run_with_progress(parent, title, status_text, task_func, *args, **kwargs):
    """
    Run a task in a background thread with a progress dialog.
    
    Args:
        parent: Parent window
        title: Dialog title
        status_text: Status message to display
        task_func: Function to run in background
        *args, **kwargs: Arguments to pass to task_func
    
    Returns: Result from task_func
    """
    result = [None]
    error = [None]
    
    def worker():
        try:
            result[0] = task_func(*args, **kwargs)
        except Exception as e:
            error[0] = e
    
    # Create and show progress dialog
    progress = ProgressDialog(parent, title)
    progress.update_status(status_text)
    
    # Start worker thread
    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    
    # Wait for thread to complete
    while thread.is_alive():
        parent.update()
        time.sleep(0.05)
    
    # Close dialog
    progress.close()
    
    # Raise error if any
    if error[0]:
        raise error[0]
    
    return result[0]
