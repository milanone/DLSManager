"""How the NSZ viewer works with and without PlotStyleKit (shared Origin style and figure editor).

The program carries the bootstrap of PlotStyleKit's integration module: when it cannot find PlotStyleKit it asks for
the folder, and without it every figure action has a fallback. The dialogs are replaced by answers; the application
runs on a hidden Tk window with synthetic samples.

    py -m unittest discover -s tests -v
"""
import json
import os
import pickle
import tempfile
import tkinter as tk
import unittest
from unittest import mock

from test_nsz_viewer import nsz, synthetic_sample          # the same module object, so patching it here works

KIT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "PlotStyleKit"))


@unittest.skipUnless(os.path.isfile(os.path.join(KIT, "plotstylekit.py")), "needs the PlotStyleKit folder next to this project")
class TestWithAndWithoutPlotStyleKit(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.root = nsz.TkinterDnD.Tk() if nsz.DND_AVAILABLE else tk.Tk()
        except tk.TclError as e:
            raise unittest.SkipTest("Tk not available: %s" % e)
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.saved = (nsz.psk, nsz.PSK_PROBLEM, nsz.PSK_CONFIG)
        nsz.PSK_CONFIG = os.path.join(self.tmp.name, "psk.json")           # never touch the real ~/.plotstylekit.json
        self.addCleanup(self._restore)
        self.app = nsz.DLSManager(self.root)
        d = synthetic_sample("a")
        self.app.samples[d["path"]] = d
        self.app.listbox.insert(tk.END, "a.nsz")
        self.app.listbox.selection_set(0)
        self.app._on_select()
        self.app.notebook.select(self.app.tab_size)
        self.root.update()

    def _restore(self):
        nsz.psk, nsz.PSK_PROBLEM, nsz.PSK_CONFIG = self.saved
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.destroy()

    def path(self, name):
        return os.path.join(self.tmp.name, name)

    def test_the_kit_is_in_use_and_styles_the_plots(self):
        self.assertIsNotNone(nsz.psk)
        ax = self.app.fig_size.axes[0]
        self.assertEqual(ax.xaxis.get_ticks_position(), 'bottom')                  # Origin: ticks on bottom/left only
        self.assertFalse(any(g.get_visible() for g in ax.xaxis.get_gridlines()))

    def test_two_panels_are_saved_at_the_double_preset(self):
        with mock.patch.object(nsz.filedialog, 'asksaveasfilename', return_value=self.path('f.fig.pickle')), \
                mock.patch.object(nsz.messagebox, 'showinfo'):
            self.app.save_figure_pickle()
        with open(self.path('f.fig.pickle'), 'rb') as f:
            fig = pickle.load(f)
        self.assertEqual(len(fig.axes), 2)
        w, h = fig.get_size_inches()
        self.assertAlmostEqual(w * 2.54, 16.0, places=3)
        self.assertAlmostEqual(h * 2.54, 9.0, places=3)

    def test_without_the_kit_the_plots_are_plain_and_the_figure_is_saved_as_it_is(self):
        nsz.psk = None
        self.app._refresh_all()
        self.assertTrue(any(g.get_visible() for g in self.app.fig_size.axes[0].xaxis.get_gridlines()))   # the dashed grid
        live = tuple(self.app.fig_size.get_size_inches())
        with mock.patch.object(nsz.filedialog, 'asksaveasfilename', return_value=self.path('plain.fig.pickle')), \
                mock.patch.object(nsz.messagebox, 'showinfo') as info, \
                mock.patch.object(nsz.messagebox, 'showerror') as err:
            self.app.save_figure_pickle()
        self.assertFalse(err.called, err.call_args)
        self.assertIn('not restyled', info.call_args[0][1])
        with open(self.path('plain.fig.pickle'), 'rb') as f:
            self.assertEqual(tuple(pickle.load(f).get_size_inches()), live)

    def test_without_the_kit_the_image_is_written_too(self):
        nsz.psk = None
        with mock.patch.object(nsz.filedialog, 'asksaveasfilename', return_value=self.path('plain.png')), \
                mock.patch.object(nsz.messagebox, 'showinfo'), mock.patch.object(nsz.messagebox, 'showerror') as err:
            self.app.save_figure_image()
        self.assertFalse(err.called, err.call_args)
        self.assertGreater(os.path.getsize(self.path('plain.png')), 5000)

    def test_without_the_kit_edit_figure_offers_to_locate_it(self):
        nsz.psk = None
        with mock.patch.object(nsz.messagebox, 'askyesno', return_value=False) as ask:
            self.app.open_figure_editor()
        self.assertTrue(ask.called)
        self.assertIsNone(nsz.psk)
        self.assertFalse(any(isinstance(w, tk.Toplevel) for w in self.root.winfo_children()))

    def test_cancelling_the_warning_is_remembered(self):
        nsz.psk = None
        with mock.patch.object(nsz.messagebox, 'askokcancel', return_value=False) as ask:
            self.app._offer_plotstylekit("PlotStyleKit was not found")
        self.assertEqual(ask.call_count, 1)
        self.assertIsNone(nsz.psk)
        with open(nsz.PSK_CONFIG, encoding='utf-8') as f:
            self.assertTrue(json.load(f)['declined'])

    def test_choosing_the_folder_starts_the_kit_and_replots(self):
        nsz.psk = None
        self.app._refresh_all()
        with mock.patch.object(nsz.messagebox, 'askokcancel', return_value=True), \
                mock.patch.object(nsz.filedialog, 'askdirectory', return_value=KIT), \
                mock.patch.object(nsz.messagebox, 'showinfo'):
            self.app._offer_plotstylekit("PlotStyleKit was not found")
        self.assertIsNotNone(nsz.psk)
        self.assertFalse(any(g.get_visible() for g in self.app.fig_size.axes[0].xaxis.get_gridlines()))   # styled now
        with open(nsz.PSK_CONFIG, encoding='utf-8') as f:
            cfg = json.load(f)
        self.assertEqual(os.path.normpath(cfg['path']), KIT)
        self.assertFalse(cfg['declined'])

    def test_a_wrong_folder_is_refused_with_the_reason(self):
        nsz.psk = None
        empty = self.path('empty')
        os.makedirs(empty)
        texts, answers = [], iter([True, False])
        with mock.patch.object(nsz.messagebox, 'askokcancel',
                               side_effect=lambda *a, **k: (texts.append(a[1]) or next(answers))), \
                mock.patch.object(nsz.filedialog, 'askdirectory', return_value=empty):
            self.app._offer_plotstylekit("PlotStyleKit was not found")
        self.assertEqual(len(texts), 2)
        self.assertIn("is not in", texts[1])
        self.assertIsNone(nsz.psk)


if __name__ == "__main__":
    unittest.main()
