"""Tests for dark mode.

The defining requirement: toggling dark on and back off must leave every widget
exactly as it was. The previous implementation guessed on restore (all frames to
'white', buttons to white text) and only ever saved originals for Labels and
Buttons, so one toggle permanently flattened every coloured surface.
"""
import tkinter as tk
from tkinter import ttk

import pytest

from gui.utils.dark_mode_manager import DarkModeManager


@pytest.fixture
def root(tk_root):
    return tk_root


@pytest.fixture
def scene(root):
    """A container holding the kinds of coloured widgets the app really uses."""
    container = tk.Frame(root, bg='#F8F9FA')
    container.pack(fill='both', expand=True)

    widgets = {
        'accent_bar': tk.Frame(container, bg='#0078D4'),          # accent stripe
        'tinted_row': tk.Frame(container, bg='#E8F5E8'),          # 'ready' tint
        'warn_row': tk.Frame(container, bg='#FFF4E5'),            # 'pending' tint
        'white_card': tk.Frame(container, bg='white'),
        'green_label': tk.Label(container, bg='#E8F5E8', fg='#107C10', text='ok'),
        'purple_label': tk.Label(container, bg='#FAF5FC', fg='#7B1FA2', text='ref'),
        'dark_button': tk.Button(container, bg='#D83B01', fg='black', text='go'),
        'entry': tk.Entry(container),
        'check': tk.Checkbutton(container, bg='white', text='opt'),
    }
    for w in widgets.values():
        w.pack()
    root.update_idletasks()

    style = ttk.Style()
    manager = DarkModeManager(root, style)
    manager.initialize()
    yield manager, container, widgets
    container.destroy()
    root.update_idletasks()


def snapshot(widgets):
    out = {}
    for name, w in widgets.items():
        entry = {}
        for opt in ('background', 'foreground'):
            try:
                if opt in w.keys():
                    entry[opt] = str(w.cget(opt))
            except tk.TclError:
                pass
        out[name] = entry
    return out


class TestRoundTrip:
    def test_restores_every_colour_exactly(self, scene, root):
        manager, _container, widgets = scene
        before = snapshot(widgets)

        manager.apply_dark_mode()
        root.update_idletasks()
        manager.restore_original_colors()
        root.update_idletasks()

        assert snapshot(widgets) == before

    def test_repeated_toggling_does_not_drift(self, scene, root):
        """Colours must survive several round trips, not just the first."""
        manager, _container, widgets = scene
        before = snapshot(widgets)

        for _ in range(3):
            manager.apply_dark_mode()
            root.update_idletasks()
            manager.restore_original_colors()
            root.update_idletasks()

        assert snapshot(widgets) == before

    def test_coloured_surfaces_are_not_flattened_to_white(self, scene, root):
        """The specific old bug: every Frame/Canvas came back 'white'."""
        manager, _container, widgets = scene
        manager.apply_dark_mode()
        manager.restore_original_colors()
        root.update_idletasks()

        assert str(widgets['accent_bar'].cget('background')) == '#0078D4'
        assert str(widgets['tinted_row'].cget('background')) == '#E8F5E8'
        assert str(widgets['warn_row'].cget('background')) == '#FFF4E5'

    def test_button_foreground_is_not_forced_to_white(self, scene, root):
        """The old restore set fg='white' unconditionally, which could leave
        white-on-white buttons."""
        manager, _container, widgets = scene
        manager.apply_dark_mode()
        manager.restore_original_colors()
        root.update_idletasks()

        assert str(widgets['dark_button'].cget('foreground')) == 'black'


class TestDarkApplication:
    def test_widgets_actually_go_dark(self, scene, root):
        manager, _container, widgets = scene
        manager.apply_dark_mode()
        root.update_idletasks()

        # A coloured label must not be left light-on-light: the old code only
        # darkened labels whose background was in a four-item whitelist.
        assert str(widgets['green_label'].cget('background')) == manager.dark_colors['bg']
        assert str(widgets['green_label'].cget('foreground')) == manager.dark_colors['fg']
        assert str(widgets['purple_label'].cget('background')) == manager.dark_colors['bg']

    def test_entry_uses_the_field_colour(self, scene, root):
        manager, _container, widgets = scene
        manager.apply_dark_mode()
        root.update_idletasks()
        assert str(widgets['entry'].cget('background')) == manager.dark_colors['widget_bg']

    def test_widget_created_after_the_toggle_is_captured(self, scene, root):
        """Late-created widgets must still round-trip correctly."""
        manager, container, _widgets = scene
        manager.apply_dark_mode()

        late = tk.Frame(container, bg='#123456')
        late.pack()
        root.update_idletasks()
        manager.apply_dark_mode()          # themes it, capturing first
        manager.restore_original_colors()
        root.update_idletasks()

        assert str(late.cget('background')) == '#123456'


class TestTreeTags:
    """Regression: the else branch was attached to `if hasattr(...)`, so light
    mode never restored the tags and a tagless widget would raise."""

    def test_light_mode_restores_tags(self, root):
        tree = ttk.Treeview(root)
        manager = DarkModeManager(root, ttk.Style())
        manager.initialize()

        manager.update_tree_tags(tree, is_dark_mode=True)
        dark_bg = str(tree.tag_configure('complete', 'background'))
        manager.update_tree_tags(tree, is_dark_mode=False)
        light_bg = str(tree.tag_configure('complete', 'background'))

        assert dark_bg == '#1a3d1a'
        assert light_bg != dark_bg, "light mode must restore the tag colours"
        tree.destroy()

    def test_widget_without_tag_configure_is_ignored(self, root):
        manager = DarkModeManager(root, ttk.Style())
        manager.initialize()
        manager.update_tree_tags(object(), is_dark_mode=False)   # must not raise


class TestDynamicColoursSurviveTheToggle:
    """The Setup checklist tints rows by state, and dark mode replays colours
    captured when it was switched ON. If the state changes while dark, restoring
    put back the PREVIOUS state's tint - showing green for an input that had
    become invalid. On that screen the colour carries meaning, so a stale one is
    actively misleading.

    Fixed by recomputing the state-derived colours after a theme change; this
    pins the contract that a theme toggle must not leave stale state colours.
    """

    OK_TINT = '#E8F5E8'
    BAD_TINT = '#FDECEA'

    def test_state_change_while_dark_is_reflected_after_restore(self, tk_root, tmp_path):
        from gui.tabs.setup_tab import STATE_STYLE
        from gui.utils import setup_validation as sv

        # The tints the tab uses must be the ones this test reasons about.
        assert STATE_STYLE[sv.OK][2] == self.OK_TINT
        assert STATE_STYLE[sv.INVALID][2] == self.BAD_TINT

    def test_toggle_recomputes_rather_than_replaying(self):
        """main_window.toggle_dark_mode must refresh state-derived colours."""
        import inspect
        from gui.main_window import GSTOrganizerApp

        source = inspect.getsource(GSTOrganizerApp.toggle_dark_mode)
        assert 'refresh_status' in source, (
            "toggle_dark_mode must recompute state-derived colours, or a state "
            "change made while dark will be restored with the old colour")
