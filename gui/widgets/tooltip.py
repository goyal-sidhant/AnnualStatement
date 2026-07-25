"""Lightweight hover tooltip, used to show full paths that are shown truncated."""
import tkinter as tk


class Tooltip:
    """Show `text` in a small popup while the pointer rests over `widget`.

    The text can be updated later via set_text() - used for path fields whose
    value changes.
    """

    def __init__(self, widget, text='', delay=450, wraplength=520):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        self._after_id = None
        self._window = None

        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def set_text(self, text):
        self.text = text or ''
        if self._window:                     # refresh while visible
            self._hide()

    # -- internals ---------------------------------------------------------
    def _schedule(self, _event=None):
        self._cancel()
        if self.text:
            self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self._window or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self._window = tk.Toplevel(self.widget)
            self._window.wm_overrideredirect(True)
            self._window.wm_geometry(f"+{x}+{y}")
            tk.Label(
                self._window, text=self.text, justify='left',
                bg='#FFFFE0', fg='#333333', relief='solid', borderwidth=1,
                font=('Segoe UI', 9), wraplength=self.wraplength,
                padx=6, pady=4,
            ).pack()
        except Exception:
            # A tooltip must never break the app
            self._window = None

    def _hide(self, _event=None):
        self._cancel()
        if self._window:
            try:
                self._window.destroy()
            except Exception:
                pass
            self._window = None
