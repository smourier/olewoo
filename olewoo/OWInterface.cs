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
    class OWInterface : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        readonly bool _topLevel;

        public OWInterface(TlibNode parent, ITypeInfo ti, TypeAttr ta, bool topLevel)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
            _topLevel = topLevel;
        }

        public override int ImageIndex => (int)ImageIndices.idx_interface;
        public override string Name => (_topLevel ? "interface " : "") + _name;
        public override string ObjectName => _name + "#i";
        public override string ShortName => _name;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override TlibNode Parent => _parent;

        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            // First add a child for every method / property, then an inherited interfaces
            // child (if applicable).

            int nfuncs = _ta.cFuncs;
            for (int idx = 0; idx < nfuncs; ++idx)
            {
                var fd = new FuncDesc(_ti, idx);
                res.Add(new OWMethod(this, _ti, fd));
            }

            if (_ta.cImplTypes > 0)
            {
                res.Add(new OWInheritedInterfaces(this, _ti, _ta));
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

            if (0 != (_ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FHIDDEN))
            {
                lprops.Add("hidden");
            }

            if (0 != (_ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FDUAL))
            {
                lprops.Add("dual");
            }

            if (0 != (_ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FRESTRICTED))
            {
                lprops.Add("restricted");
            }

            if (0 != (_ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE))
            {
                lprops.Add("nonextensible");
            }

            if (0 != (_ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FOLEAUTOMATION))
            {
                lprops.Add("oleautomation");
            }

            for (int i = 0; i < lprops.Count; ++i)
            {
                ih.AppendLine("  " + lprops[i] + (i < (lprops.Count - 1) ? "," : ""));
            }
            ih.AppendLine("]");

            if (_ta.cImplTypes > 0)
            {
                _ti.GetRefTypeOfImplType(0, out int href);
                _ti.GetRefTypeInfo(href, out ITypeInfo ti2);
                ih.AddString("interface " + _name + " : ");
                ih.AddLink(ti2.GetName(), "i");
                ih.AppendLine(" {");
            }
            else
            {
                ih.AppendLine("interface " + _name + " {");
            }

            using (new IDLHelperTab(ih))
            {
                Children.ForEach(x => x.BuildIDLInto(ih));
            }
            ih.AppendLine("};");
        }
    }
}

