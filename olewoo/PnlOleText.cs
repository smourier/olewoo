/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace olewoo
{
    public partial class PnlOleText : UserControl
    {
        LinkedList<TreeNode> _history;
        RichIDLFormatter _fo;
        SetTextDelg _std; // for setting the title of the tab I'm contained in.
        PnlTextOrTabbed _parent;
        bool _rewinding;
        TreeNode _n;

        public PnlOleText()
        {
            InitializeComponent();
            rtfOleText.SetTabs(4);
            _fo = new RichIDLFormatter(rtfOleText);
            _history = new LinkedList<TreeNode>();
            rtfOleText.LinkClicked += LinkClicked;
            rtfOleText.KeyDown += TxtOleText_KeyDown;
            rtfOleText.KeyPress += TxtOleText_KeyPress;
        }

        public PnlOleText(SetTextDelg s) : this() // makes the c++ me weep.
        {
            _std = s;
        }

        private void AddHistory(TreeNode tn)
        {
            if (_rewinding) return;
            if (_history.Count > 50) _history.RemoveFirst(); // O(N)? :P
            _history.AddLast(tn);
        }

        public void RewindOne()
        {
            if (_history.Count == 0)
                return; // O(N)? :P

            try
            {
                _rewinding = true;
                TreeNode tn = _history.Last.Value;
                if (tn != null)
                {
                    tn.TreeView.SelectedNode = tn;
                }

                _history.RemoveLast();
            }
            finally
            {
                _rewinding = false;
            }

        }

        void LinkClicked(object sender, LinkClickedEventArgs e)
        {
            NamedNode nn = _parent.NodeLocator.FindLinkMatch(e.LinkText);
            if (nn != null)
            {
                nn.TreeNode.TreeView.SelectedNode = nn.TreeNode;
            }
        }

        public PnlTextOrTabbed TabParent
        {
            set => _parent = value;
        }

        public TreeNode TreeNode
        {
            get => _n;
            set
            {
                using (new TBUpdateSuspender(rtfOleText))
                {
                    AddHistory(_n);
                    TlibNode tn = (value == null) ? null : (value.Tag as TlibNode);
                    _n = value;
                    rtfOleText.Text = "";
                    if (tn != null)
                    {
                        tn.BuildIDLInto(_fo);
                        _fo.Flush();
                        rtfOleText.Select(0, 0);
                    }

                    if (_std != null)
                    {
                        string sn = tn == null ? "..." : tn.ShortName;
                        if (sn == null) sn = tn.Name;
                        if (sn.Length > 10) sn = sn.Substring(0, 9) + "...";
                        _std(sn);
                    }
                }
            }
        }

        public bool FindNextText(string needle, bool searchDown)
        {
            if (rtfOleText.Text == "")
                return false;

            int idx = rtfOleText.SelectionStart;
            int pos = -1;
            if (searchDown)
            {
                try
                {
                    pos = rtfOleText.Text.IndexOf(needle, idx + 1, StringComparison.CurrentCultureIgnoreCase);
                }
                catch (IndexOutOfRangeException)
                {
                }
                if (pos == -1)
                {
                    pos = rtfOleText.Text.IndexOf(needle, 0, StringComparison.CurrentCultureIgnoreCase);
                }
            }
            else
            {
                // Something is broken here - when debugging, the 1 becomes an invalid number!
                // Search up disabled to stop this occurrence.
                throw new NotImplementedException();
                //                pos = txtOleText.Text.LastIndexOf(needle, 1, idx, StringComparison.CurrentCultureIgnoreCase);
            }

            if (pos != -1)
            {
                rtfOleText.SelectionStart = pos;
                rtfOleText.SelectionLength = needle.Length;
                rtfOleText.ScrollToCaret();
                return true;
            }
            return false;
        }

        public void TxtOleText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                switch (e.KeyCode)
                {
                    case Keys.F:
                        var fd = new FindDialog(this);
                        fd.ShowDialog(rtfOleText);
                        break;

                    case Keys.A:
                        rtfOleText.SelectAll();
                        break;

                    case Keys.C:
                        Clipboard.SetText(rtfOleText.SelectedText);
                        break;
                }
            }
        }

        public void TxtOleText_KeyPress(object sender, KeyPressEventArgs e) => e.Handled = true;

        private void TxtOleText_TextChanged(object sender, EventArgs e)
        {

        }
    }

    public delegate void SetTextDelg(string s);

    class RichIDLFormatter : IDLFormatter
    {
        StringBuilder _sb;
        RichTextBoxLinks.RichTextBoxEx _rtb;
        bool _bPendingApplyTabs;

        public RichIDLFormatter(RichTextBoxLinks.RichTextBoxEx rtb)
        {
            _sb = new StringBuilder();
            _rtb = rtb;
            _bPendingApplyTabs = false;
        }

        public override string ToString() => _sb.ToString();

        public override void NewLine()
        {
            _sb.Append("\r\n");
            _bPendingApplyTabs = true;
        }

        private void ApplyTabs()
        {
            _bPendingApplyTabs = false;
            if (_tabdepth > 0)
            {
                string s = "";
                for (int x = 0; x < _tabdepth; ++x)
                {
                    s += "\t";
                }
                _sb.Append(s);
            }
        }

        public override void AddString(string s)
        {
            if (_bPendingApplyTabs) ApplyTabs();
            _sb.Append(s);
        }

        public override void AddLink(string s, string s2)
        {
            if (_bPendingApplyTabs) ApplyTabs();
            Flush();
            _rtb.InsertLink(s, s2);
        }

        public void Flush()
        {
            _rtb.AppendText(_sb.ToString());
            _sb.Length = 0;
        }
    }
}
