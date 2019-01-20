/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace olewoo
{
    public partial class wooctrl : UserControl
    {
        enum SortType
        {
            Sorted_Numerically,
            Sorted_AlphaUp,
            Sorted_AlphaDown,
            Sorted_Max
        }

        OWTypeLib _tlib;
        NodeLocator _nl;
        ImageList _iml;
        SortType _sort;

        public wooctrl(ImageList imglstTreeNodes, ImageList imglstMisc, OWTypeLib tlib)
        {
            _tlib = tlib;
            _nl = new NodeLocator();
            _iml = imglstMisc;
            _sort = SortType.Sorted_Numerically;

            InitializeComponent();
            txtOleDescrPlain.ParentCtrl = this;
            tvLibDisp.ImageList = imglstTreeNodes;
            Dock = DockStyle.Fill;

            tvLibDisp.Nodes.Add(GenNodeTree(tlib, _nl));
            txtOleDescrPlain.NodeLocator = _nl;
            tvLibDisp.Nodes[0].Expand();
        }

        public ImageList ImageList => _iml;

        // Note that this generates redundant tree nodes, i.e. many definitions of (eg) IUnknown
        // this is because Trees don't have support for child sharing. (how would the parent property work? :)
        private TreeNode GenNodeTree(TlibNode tln, NodeLocator nl)
        {
            var children = tln.Children.ConvertAll(x => GenNodeTree(x, nl)).ToArray();
            var tn = new TreeNode(tln.Name, tln.ImageIndex, (int)TlibNode.ImageIndices.idx_selected, children);
            tn.Tag = tln;
            nl.Add(tn);
            return tn;
        }

        private void ClearMatches()
        {
            pnlMatchesList.Visible = false;
            lstNodeMatches.Items.Clear();
        }

        // Search through the registered names for the tree nodes.
        // When we hit one, select that node.
        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            string text = txtSearch.Text;
            if (text == "")
            {
                ClearMatches();
            }
            else
            {
                try
                {
                    List<NamedNode> tn = _nl.FindMatches(text);
                    lstNodeMatches.Items.Clear();
                    if (tn != null && tn.Count > 0)
                    {
                        tvLibDisp.ActivateNode(tn.First().TreeNode);
                        lstNodeMatches.Items.AddRange(tn.ToArray());
                    }
                    pnlMatchesList.Visible = true;
                }
                catch (Exception)
                {
                }
            }
        }

        private void LstNodeMatches_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!(lstNodeMatches.SelectedItem is NamedNode nn))
                return;

            tvLibDisp.ActivateNode(nn.TreeNode);
        }

        private void TvLibDisp_AfterSelect(object sender, TreeViewEventArgs e) => txtOleDescrPlain.SetCurrentNode(e.Node);
        private void BtnHideMatches_Click(object sender, EventArgs e) => txtSearch.Text = string.Empty;
        private void BtnAddNodeTabl_Click(object sender, EventArgs e) => txtOleDescrPlain.AddTab(tvLibDisp.SelectedNode);
        public void SelectTreeNode(TreeNode tn) => tvLibDisp.ActivateNode(tn);

        delegate int CompDelg(TlibNode x, TlibNode y);

        class OleNodeComparer : IComparer<TreeNode>
        {
            readonly CompDelg _cd;

            public OleNodeComparer(SortType t)
            {
                switch (t)
                {
                    case SortType.Sorted_Numerically:
                    default: // shouldn't happen...
                        _cd = CompareNum;
                        break;
                    case SortType.Sorted_AlphaUp:
                        _cd = CompareAlphaUp;
                        break;
                    case SortType.Sorted_AlphaDown:
                        _cd = CompareAlphaDown;
                        break;
                }
            }

            public int Compare(TreeNode x, TreeNode y) => _cd(x.Tag as TlibNode, y.Tag as TlibNode);
            int CompareNum(TlibNode x, TlibNode y) => x.Idx.CompareTo(y.Idx);
            int CompareAlphaUp(TlibNode x, TlibNode y) => x.Name.CompareTo(y.Name);
            int CompareAlphaDown(TlibNode x, TlibNode y) => y.Name.CompareTo(x.Name);
        }

        private void BtnSortAlpha_Click(object sender, EventArgs e)
        {
            using (new UpdateSuspender(tvLibDisp))
            {
                _sort++;
                if (_sort == SortType.Sorted_Max)
                {
                    _sort = SortType.Sorted_Numerically;
                }

                var nlst = new List<TreeNode>();
                TreeNode root = tvLibDisp.Nodes[0];
                var enm = root.Nodes.GetEnumerator();
                while (enm.MoveNext())
                {
                    nlst.Add(enm.Current as TreeNode);
                }

                nlst.Sort(new OleNodeComparer(_sort));
                root.Nodes.Clear();
                root.Nodes.AddRange(nlst.ToArray());
            }
        }

        private void BtnRewind_Click(object sender, EventArgs e) => txtOleDescrPlain.RewindOne();

    }

    static class Xtras
    {
        public static void ActivateNode(this TreeView tv, TreeNode tn)
        {
            if (tn != null)
            {
                tv.SelectedNode = tn;
                tn.EnsureVisible();
            }
        }
    }

    public class NamedNode
    {
        private readonly TreeNode _tn;
        private TlibNode _tln;

        public NamedNode(TreeNode tn)
        {
            _tn = tn;
            _tln = tn.Tag as TlibNode;
        }

        public string Name => _tln.ShortName;
        public string ObjectName => _tln.ObjectName;
        public TreeNode TreeNode => _tn;
        public override string ToString() => _tln.ShortName;
    }

    public class NodeLocator
    {
        List<NamedNode> _nodes;
        Dictionary<string, NamedNode> _linkmap;

        public NodeLocator()
        {
            _nodes = new List<NamedNode>();
            _linkmap = new Dictionary<string, NamedNode>();
        }

        public void Add(TreeNode tn)
        {
            var tli = tn.Tag as TlibNode;
            string name = tli.ShortName;
            if (name != null)
            {
                var nn = new NamedNode(tn);
                _nodes.Add(nn);
                string oname = tli.ObjectName;
                if (!string.IsNullOrEmpty(oname) && !_linkmap.ContainsKey(oname))
                {
                    _linkmap[oname] = nn;
                }
            }
        }
        // O(N).  FIX!
        public List<NamedNode> FindMatches(string text)
        {
            var re = new Regex("^.*" + text, RegexOptions.IgnoreCase);
            return _nodes.FindAll(x => re.IsMatch(x.Name));
        }

        public NamedNode FindLinkMatch(string text)
        {
            if (_linkmap.ContainsKey(text))
                return _linkmap[text];

            return null;
        }
    }
}
