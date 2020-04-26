/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using olewoo.interop;

namespace olewoo
{
    class OWDispInterface : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        TypeAttr _ta;
        ITypeInfo _ti;
        readonly bool _topLevel;
        OWIDispatchMethods _methodChildren;
        OWIDispatchProperties _propChildren;

        public OWDispInterface(TlibNode parent, ITypeInfo ti, TypeAttr ta, bool topLevel)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
            _topLevel = topLevel;
        }

        public override string Name => (_topLevel ? "dispinterface " : "") + _name;
        public override string ShortName => _name;
        public override string ObjectName => _name + "#di";
        public override string ToString() => Name;

        // Don't show a dispinterface at top level, UNLESS the corresponding interface is not itself at top level. 
        public override int ImageIndex => (int)ImageIndices.idx_dispinterface;
        public override TlibNode Parent => _parent;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => !(interfaceNames.Contains(ShortName));
        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            if (_ta.cVars > 0)
            {
                _propChildren = new OWIDispatchProperties(this);
                res.Add(_propChildren);
            }

            if (_ta.cFuncs > 0)
            {
                _methodChildren = new OWIDispatchMethods(this);
                res.Add(_methodChildren);
            }

            if (_ta.cImplTypes > 0)
            {
                res.Add(new OWDispInterfaceInheritedInterfaces(this, _ti, _ta));
            }
            return res;
        }

        public List<TlibNode> MethodChildren()
        {
            var res = new List<TlibNode>();
            int nfuncs = _ta.cFuncs;
            for (int idx = 0; idx < nfuncs; ++idx)
            {
                var fd = new FuncDesc(_ti, idx);
                //                if (0 == (fd.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED))
                res.Add(new OWMethod(this, _ti, fd));
            }
            return res;
        }

        public List<TlibNode> PropertyChildren()
        {
            var res = new List<TlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                var vd = new VarDesc(_ti, x);
                res.Add(new OWDispProperty(this, _ti, vd));
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
            ih.AppendLine("dispinterface " + _name + " {");

            if (_ta.cFuncs > 0 || _ta.cVars > 0)
            {
                // Naughty, but rely on side effect of verifying children.
                var children = Children;
                using (new IDLHelperTab(ih))
                {
                    if (_propChildren != null)
                    {
                        _propChildren.BuildIDLInto(ih);
                    }
                    if (_methodChildren != null)
                    {
                        _methodChildren.BuildIDLInto(ih);
                    }
                }
            }
            ih.AppendLine("};");
        }
    }
}

