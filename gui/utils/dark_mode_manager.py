# gui/utils/dark_mode_manager.py
"""Dark mode with real state restoration.

The previous implementation mutated widgets and then tried to reverse that by
GUESSING: every Frame and Canvas was forced to 'white' on restore, buttons were
forced to white text, and originals were only ever kept for Labels and Buttons.
One toggle therefore flattened every coloured surface in the app permanently -
accent bars, section headers, the Setup checklist's state tints - until restart.

This version captures each widget's actual colours before touching it and
replays exactly those values on restore. It is also generic: instead of a
per-class if/elif chain that silently skipped anything unfamiliar, it asks each
widget which colour options it supports.
"""

import logging
import tkinter as tk
import weakref

from utils.constants import GUI_CONFIG

logger = logging.getLogger(__name__)

# Colour options worth theming, if the widget supports them.
THEMED_OPTIONS = (
    'background', 'foreground', 'insertbackground', 'selectcolor',
    'activebackground', 'activeforeground', 'highlightbackground',
    'troughcolor', 'selectbackground', 'selectforeground',
)

# Widgets that read as an input field rather than a surface.
FIELD_CLASSES = {'Entry', 'Text', 'Spinbox', 'Listbox', 'TCombobox'}
BUTTON_CLASSES = {'Button', 'Checkbutton', 'Radiobutton', 'Menubutton'}

# ttk styles to save and re-theme. The old version covered only four, so ttk
# buttons and scrollbars stayed bright in dark mode.
TTK_STYLES = {
    'TFrame': ('background',),
    'TLabel': ('background', 'foreground'),
    'TNotebook': ('background',),
    'TNotebook.Tab': ('background', 'foreground'),
    'TButton': ('background', 'foreground'),
    'TEntry': ('fieldbackground', 'foreground'),
    'TCombobox': ('fieldbackground', 'foreground'),
    'TScrollbar': ('background', 'troughcolor'),
    'TProgressbar': ('background', 'troughcolor'),
    'Treeview': ('background', 'foreground', 'fieldbackground'),
    'Treeview.Heading': ('background', 'foreground'),
}


class DarkModeManager:
    """Toggle dark mode, restoring the exact colours that were there before."""

    def __init__(self, root, style):
        self.root = root
        self.style = style
        self.is_initialized = False
        self.is_dark_mode = False

        # widget -> {option: original value}
        #
        # Keyed by the WIDGET OBJECT, not its path string: Tk reuses path names
        # after a widget is destroyed, and this app rebuilds rows regularly
        # (display_clients, update_client_tree). Keying by path meant a
        # destroyed Frame's colours could be replayed onto a Text created at the
        # same path. Weak keys so destroyed widgets do not accumulate.
        self._original = weakref.WeakKeyDictionary()
        self._original_ttk = {}
        self._original_root_bg = None

        self.dark_colors = {
            'bg': '#2b2b2b',
            'fg': '#ffffff',
            'widget_bg': '#3b3b3b',
            'button_bg': '#404040',
            'trough': '#333333',
        }

    # ------------------------------------------------------------------ setup
    def initialize(self):
        if not self.is_initialized:
            self._original_root_bg = self.root.cget('bg')
            self._capture_ttk()
            self.is_initialized = True

    def _capture_ttk(self):
        for style_name, options in TTK_STYLES.items():
            saved = {}
            for option in options:
                try:
                    value = self.style.lookup(style_name, option)
                except Exception:
                    continue
                # Record the value even when it is EMPTY. An empty lookup means
                # the theme leaves that option unset - and skipping those was a
                # real bug: Treeview.fieldbackground starts unset, so after dark
                # mode set it there was nothing recorded to undo it and the
                # client list stayed dark grey in light mode. Restoring it to ''
                # returns it to unset.
                saved[option] = value if value is not None else ''
            self._original_ttk[style_name] = saved

    # ------------------------------------------------------------- public API
    def apply_dark_mode(self):
        self.is_dark_mode = True

        # TWO PASSES, and the order matters. Capture the whole tree BEFORE
        # anything is mutated: doing it in one pass meant the root window (and
        # anything else re-coloured on the way in) was recorded with its dark
        # value and could never be restored.
        self._walk(self.root, self._capture_widget)

        self._apply_ttk_dark()
        self._walk(self.root, self._to_dark)

    def restore_original_colors(self):
        self.is_dark_mode = False
        self._restore_ttk()
        self._walk(self.root, self._to_original)

    # ------------------------------------------------------------------- ttk
    def _apply_ttk_dark(self):
        d = self.dark_colors
        dark = {
            'TFrame': {'background': d['bg']},
            'TLabel': {'background': d['bg'], 'foreground': d['fg']},
            'TNotebook': {'background': d['bg']},
            'TNotebook.Tab': {'background': d['widget_bg'], 'foreground': d['fg']},
            'TButton': {'background': d['button_bg'], 'foreground': d['fg']},
            'TEntry': {'fieldbackground': d['widget_bg'], 'foreground': d['fg']},
            'TCombobox': {'fieldbackground': d['widget_bg'], 'foreground': d['fg']},
            'TScrollbar': {'background': d['button_bg'], 'troughcolor': d['trough']},
            'TProgressbar': {'background': GUI_CONFIG['colors']['primary'],
                             'troughcolor': d['trough']},
            'Treeview': {'background': d['widget_bg'], 'foreground': d['fg'],
                         'fieldbackground': d['widget_bg']},
            'Treeview.Heading': {'background': d['button_bg'], 'foreground': d['fg']},
        }
        for style_name, options in dark.items():
            try:
                self.style.configure(style_name, **options)
            except Exception as e:
                logger.debug(f"Could not theme {style_name}: {e}")

        try:
            self.style.map('TNotebook.Tab',
                           background=[('selected', d['button_bg'])],
                           foreground=[('selected', d['fg'])])
        except Exception:
            pass

    def _restore_ttk(self):
        for style_name, options in self._original_ttk.items():
            if not options:
                continue
            try:
                # Options captured as '' are re-applied as '' on purpose, which
                # clears the value dark mode set and returns them to unset.
                self.style.configure(style_name, **options)
            except Exception as e:
                logger.debug(f"Could not restore {style_name}: {e}")

        try:
            self.style.map('TNotebook.Tab',
                           background=[('selected', '#f0f0f0')],
                           foreground=[('selected', 'black')])
            self.style.map('Treeview',
                           foreground=[('selected', 'black')],
                           background=[('selected', GUI_CONFIG['colors']['primary'])])
        except Exception:
            pass

    # --------------------------------------------------------------- widgets
    def _walk(self, widget, action):
        """Depth-first over the whole widget tree, skipping nothing silently."""
        try:
            action(widget)
        except tk.TclError:
            pass                      # widget went away mid-walk
        except Exception as e:
            logger.debug(f"Theming {widget} failed: {e}")

        try:
            children = widget.winfo_children()
        except Exception:
            return
        for child in children:
            self._walk(child, action)

    def _supported(self, widget):
        """The themed options this particular widget actually accepts."""
        try:
            keys = set(widget.keys())
        except Exception:
            return ()
        return tuple(o for o in THEMED_OPTIONS if o in keys)

    def _capture(self, widget, options):
        """Save the widget's current colours once, before anything changes them."""
        try:
            if widget in self._original:
                return
        except TypeError:
            return                      # not weak-referenceable
        saved = {}
        for option in options:
            try:
                saved[option] = widget.cget(option)
            except Exception:
                pass
        try:
            self._original[widget] = saved
        except TypeError:
            pass

    def _capture_widget(self, widget):
        """Capture pass: record this widget's colours as they are right now."""
        options = self._supported(widget)
        if options:
            self._capture(widget, options)

    def _to_dark(self, widget):
        options = self._supported(widget)
        if not options:
            return
        # Belt and braces: the capture pass has already run, but a widget created
        # after the first toggle would otherwise be mutated before being saved.
        self._capture(widget, options)

        cls = widget.winfo_class()
        d = self.dark_colors
        surface = d['widget_bg'] if cls in FIELD_CLASSES else (
            d['button_bg'] if cls in BUTTON_CLASSES else d['bg'])

        values = {}
        if 'background' in options:
            values['background'] = surface
        if 'foreground' in options:
            values['foreground'] = d['fg']
        if 'insertbackground' in options:
            values['insertbackground'] = d['fg']
        if 'selectcolor' in options:
            values['selectcolor'] = d['widget_bg']
        if 'activebackground' in options:
            values['activebackground'] = d['widget_bg']
        if 'activeforeground' in options:
            values['activeforeground'] = d['fg']
        if 'troughcolor' in options:
            values['troughcolor'] = d['trough']

        if values:
            widget.configure(**values)

    def _to_original(self, widget):
        """Replay exactly what was captured. No guessing, no defaults."""
        try:
            saved = self._original.get(widget)
        except TypeError:
            saved = None
        if not saved:
            return
        try:
            widget.configure(**saved)
        except tk.TclError:
            pass
        except Exception as e:
            logger.debug(f"Restoring {widget} failed: {e}")

    # ------------------------------------------------------------------ trees
    def update_tree_tags(self, tree_widget, is_dark_mode=None):
        """Recolour a Treeview's complete/incomplete row tags.

        The previous version attached its `else` to `if hasattr(...)` instead of
        `if is_dark_mode`, so light mode never restored the tags (rows stayed
        dark) and a widget without tag_configure would raise AttributeError.
        """
        if is_dark_mode is None:
            is_dark_mode = self.is_dark_mode
        if not hasattr(tree_widget, 'tag_configure'):
            return

        if is_dark_mode:
            tree_widget.tag_configure('complete', background='#1a3d1a',
                                      foreground=self.dark_colors['fg'])
            tree_widget.tag_configure('incomplete', background='#3d2f1a',
                                      foreground=self.dark_colors['fg'])
        else:
            tree_widget.tag_configure('complete',
                                      background=GUI_CONFIG['colors']['complete'],
                                      foreground='black')
            tree_widget.tag_configure('incomplete',
                                      background=GUI_CONFIG['colors']['incomplete'],
                                      foreground='black')
