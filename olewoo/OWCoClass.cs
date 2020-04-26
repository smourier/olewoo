/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using olewoo.interop;

namespace olewoo
{
    class OWCoClass : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        TypeAttr _ta;
        ITypeInfo _ti;

        public OWCoClass(TlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
        }

        public override string Name => "coclass " + _name;
        public override string ObjectName => _name + "#c";
        public override string ShortName => _name;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override int ImageIndex => (int)ImageIndices.idx_coclass;
        public override TlibNode Parent => _parent;
        public override string ToString() => Name;

        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            for (int x = 0; x < _ta.cImplTypes; ++x)
            {
                _ti.GetRefTypeOfImplType(x, out int href);
                _ti.GetRefTypeInfo(href, out ITypeInfo ti2);
                CommonBuildTlibNode(this, ti2, false, false, res);
            }
            return res;
        }

        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("[");
            var lprops = new List<string>();
            lprops.Add("uuid(" + _ta.guid + ")");
            string help = _ti.GetHelpDocumentationById(-1, out int context);
            AddHelpStringAndContext(lprops, help, context);
            for (int i = 0; i < lprops.Count; ++i)
            {
                ih.AppendLine("  " + lprops[i] + (i < (lprops.Count - 1) ? "," : ""));
            }
            ih.AppendLine("]");
            ih.AppendLine("coclass " + _name + " {");
            using (new IDLHelperTab(ih))
            {
                for (int x = 0; x < _ta.cImplTypes; ++x)
                {
                    _ti.GetRefTypeOfImplType(x, out int href);
                    _ti.GetRefTypeInfo(href, out ITypeInfo ti2);
                    _ti.GetImplTypeFlags(x, out IMPLTYPEFLAGS itypflags);
                    var res = new List<string>();
                    if (0 != (itypflags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT))
                    {
                        res.Add("default");
                    }

                    if (0 != (itypflags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE))
                    {
                        res.Add("source");
                    }

                    if (res.Count > 0)
                    {
                        ih.AddString("[" + string.Join(", ", res.ToArray()) + "] ");
                    }

                    ih.AddString("interface ");
                    ih.AddLink(ti2.GetName(), "i");
                    ih.AppendLine(";");
                }
            }
            ih.AppendLine("};");
        }
    }
}