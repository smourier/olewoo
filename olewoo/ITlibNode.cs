﻿/**************************************
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
    public abstract class ITlibNode
    {
        public delegate List<ITlibNode> dlgCreateChildren();

        public enum ImageIndices
        {
            idx_coclass,
            idx_const,
            idx_dispinterface,
            idx_enum,
            idx_interface,
            idx_strucmem,
            idx_struct,
            idx_typelib,
            idx_methodlist,
            idx_method,
            idx_typedef,
            idx_module,
            idx_constlist,
            idx_selected,
            idx_propertylist
        };

        private List<ITlibNode> _children;

        public abstract ITlibNode Parent { get; }
        public abstract List<ITlibNode> GenChildren();
        public abstract string Name { get; }
        public int Idx { get; set; }
        public abstract string ShortName { get; }
        public abstract string ObjectName { get; }
        public abstract int ImageIndex { get; }

        public List<ITlibNode> Children
        {
            get
            {
                if (_children == null)
                {
                    _children = GenChildren();
                    for (int i = 0; i < _children.Count; ++i)
                    {
                        _children[i].Idx = i;
                    }
                }
                return _children;
            }
        }

        public abstract bool DisplayAtTLBLevel(ICollection<string> interfaceNames);

        public void CommonBuildTlibNode(ITlibNode parent, ITypeInfo ti, bool topLevel, bool swapfordispatch, List<ITlibNode> res)
        {
            TypeAttr ta = ti.GetTypeAttr();
            switch (ta.typekind)
            {
                case TYPEKIND.TKIND_DISPATCH:
                    res.Add(new OWDispInterface(this, ti, ta, topLevel));
                    if (swapfordispatch && ITypeInfoXtra.SwapForInterface(ref ti, ref ta))
                    {
                        res.Add(new OWInterface(this, ti, ta, topLevel));
                    }
                    break;

                case TYPEKIND.TKIND_INTERFACE:
                    res.Add(new OWInterface(this, ti, ta, topLevel));
                    break;

                case TYPEKIND.TKIND_ALIAS:
                    res.Add(new OWTypeDef(this, ti, ta));
                    break;

                case TYPEKIND.TKIND_ENUM:
                    res.Add(new OWEnum(this, ti, ta));
                    break;

                case TYPEKIND.TKIND_COCLASS:
                    res.Add(new OWCoClass(this, ti, ta));
                    break;

                case TYPEKIND.TKIND_RECORD:
                    res.Add(new OWRecord(this, ti, ta));
                    break;

                case TYPEKIND.TKIND_MODULE:
                    res.Add(new OWModule(this, ti, ta));
                    break;
            }
        }

        public abstract void BuildIDLInto(IDLFormatter ih);

        protected void AddHelpStringAndContext(List<string> lprops, string help, int context)
        {
            if (!string.IsNullOrEmpty(help))
            {
                lprops.Add("helpstring(\"" + help + "\")");
            }

            if (context != 0)
            {
                lprops.Add("helpcontext(" + context.PaddedHex() + ")");
            }
        }
    }
}

