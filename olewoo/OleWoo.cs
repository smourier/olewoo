using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace olewoo
{
    public partial class OleWoo : Form
    {
        private MRUList _mrufiles;

        public OleWoo()
        {
            InitializeComponent();
            tcTypeLibs.ImageList = imgListMisc;
            _mrufiles = new MRUList(@"Software\benf.org\olewoo\MRU");
            var args = Environment.GetCommandLineArgs().ToList();
            args.RemoveAt(0);
            foreach (string arg in args)
            {
                OpenFile(System.IO.Path.GetFullPath(arg));
            }
        }

        private void OpenFile(string fname)
        {
            var tl = new OWTypeLib(fname);
            var tp = new TabPage(tl.ShortName);
            tp.ImageIndex = 0;
            var wc = new wooctrl(imglstTreeNodes, imgListMisc, tl);
            tp.Controls.Add(wc);
            tp.Tag = tl;
            wc.Dock = DockStyle.Fill;
            tcTypeLibs.TabPages.Add(tp);
            _mrufiles.AddItem(fname);
            _mrufiles.Flush();
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Supported Files (*.dll;*.tlb;*.exe;*.ocx)|*.dll;*.tlb;*.exe;*.ocx|Dll files (*.dll)|*.dll|Typelibraries (*.tlb)|*.tlb|Executables (*.exe)|*.exe|ActiveX controls (*.ocx)|*.ocx|All files (*.*)|*.*";
            ofd.CheckFileExists = true;
            switch (ofd.ShowDialog(this))
            {
                case DialogResult.OK:
                    OpenFile(ofd.FileName);
                    break;
            }
        }

        private void AboutOleWooToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var ab = new AboutBox();
            ab.ShowDialog();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) => Application.Exit();

        private void OpenMRUItem_Click(object sender, EventArgs e)
        {
            if (!(sender is ToolStripMenuItem mni))
                return;

            OpenFile(mni.Tag as string);
        }

        private void ClearMRUItem_Click(object sender, EventArgs e)
        {
            _mrufiles.Clear();
            _mrufiles.Flush();
        }

        delegate void VoidDelg();

        private void FileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /* Dynamically populate the file item with the MRU */
            fileToolStripMenuItem.DropDownItems.Clear();

            var tsis = new List<ToolStripItem>();
            var tmiOpen = new ToolStripMenuItem();
            tmiOpen.Name = "openToolStripMenuItem";
            tmiOpen.ShortcutKeys = (Keys.Control | Keys.O);
            tmiOpen.Size = new Size(208, 22);
            tmiOpen.Text = "&Open Typelibrary";
            tmiOpen.Click += new EventHandler(OpenToolStripMenuItem_Click);
            tsis.Add(tmiOpen);

            void addSep()
            {
                var tmiSep = new ToolStripSeparator();
                tmiSep.Name = "toolStripMenuItem1";
                tmiSep.Size = new Size(205, 6);
                tsis.Add(tmiSep);
            }

            addSep();

            string[] mru = _mrufiles.Items;
            if (mru.Length > 0)
            {
                int idx = 1;
                foreach (string mrui in mru)
                {
                    var tmiMru = new ToolStripMenuItem();
                    tmiMru.Tag = mrui;
                    tmiMru.Size = new Size(208, 22);
                    string label = mrui;
                    if (label.Length > 35) label = label.Substring(0, 10) + "..." + label.Substring(label.Length - 20);
                    tmiMru.Text = "&" + (idx++) + " " + label;
                    tmiMru.Click += new EventHandler(OpenMRUItem_Click);
                    tsis.Add(tmiMru);
                }
                addSep();
                {
                    var tmiMru = new ToolStripMenuItem();
                    tmiMru.Size = new Size(208, 22);
                    tmiMru.Text = "&Clear Recent items list";
                    tmiMru.Click += new EventHandler(ClearMRUItem_Click);
                    tsis.Add(tmiMru);
                }
                addSep();
            }

            var tmiExit = new ToolStripMenuItem();
            tmiExit.Name = "exitToolStripMenuItem";
            tmiExit.Size = new Size(208, 22);
            tmiExit.Text = "E&xit";
            tmiExit.Click += new EventHandler(ExitToolStripMenuItem_Click);
            tsis.Add(tmiExit);

            fileToolStripMenuItem.DropDownItems.AddRange(tsis.ToArray());
        }

        private void RegisterContextHandlerToolStripMenuItem_Click(object sender, EventArgs e) => FileShellExtension.Register("dllfile", "OleWoo Context Menu", "Open with OleWoo", string.Format("\"{0}\" \"%L\"", Application.ExecutablePath));
        private void UnregisterContextHandlerToolStripMenuItem_Click(object sender, EventArgs e) => FileShellExtension.Unregister("dllfile", "OleWoo Context Menu");
    }
}
