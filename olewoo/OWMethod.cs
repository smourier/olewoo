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
    class OWMethod : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        FuncDesc _fd;
        ITypeInfo _ti;

        public OWMethod(TlibNode parent, ITypeInfo ti, FuncDesc fd)
        {
            _parent = parent;
            _ti = ti;
            _fd = fd;

            string[] names = fd.GetNames(ti);
            string functionname = names[0];

            _name = names[0];
        }

        public override string Name => _name;
        public override string ShortName => _name;
        public override string ObjectName => _name + "#m";
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override int ImageIndex => (int)ImageIndices.idx_method;
        public override TlibNode Parent => _parent;

        private string ParamFlagsDescription(ParamDesc pd)
        {
            PARAMFLAG flg = pd.wParamFlags;
            var res = new List<string>();
            if (0 != (flg & PARAMFLAG.PARAMFLAG_FIN))
            {
                res.Add("in");
            }

            if (0 != (flg & PARAMFLAG.PARAMFLAG_FOUT))
            {
                res.Add("out");
            }

            if (0 != (flg & PARAMFLAG.PARAMFLAG_FRETVAL))
            {
                res.Add("retval");
            }

            if (0 != (flg & PARAMFLAG.PARAMFLAG_FOPT))
            {
                res.Add("optional");
            }

            if (0 != (flg & PARAMFLAG.PARAMFLAG_FHASDEFAULT))
            {
                res.Add("defaultvalue(" + ITypeInfoXtra.QuoteString(pd.varDefaultValue) + ")");
            }

            return "[" + string.Join(", ", res.ToArray()) + "]";
        }

        public bool Property => _fd.invkind != INVOKEKIND.INVOKE_FUNC;

        delegate void GenPt(int x);

        public override List<TlibNode> GenChildren() => new List<TlibNode>();
        public override void BuildIDLInto(IDLFormatter ih) => BuildIDLInto(ih, false);
        public void BuildIDLInto(IDLFormatter ih, bool bAsDispatch)
        {
            bool memIdInSpecialRange = (_fd.memid >= 0x60000000 && _fd.memid < 0x60020000);
            var lprops = new List<string>();
            if (!memIdInSpecialRange)
            {
                lprops.Add("id(" + _fd.memid.PaddedHex() + ")");
            }

            switch (_fd.invkind)
            {
                case INVOKEKIND.INVOKE_PROPERTYGET:
                    lprops.Add("propget");
                    break;
                case INVOKEKIND.INVOKE_PROPERTYPUT:
                    lprops.Add("propput");
                    break;
                case INVOKEKIND.INVOKE_PROPERTYPUTREF:
                    lprops.Add("propputref");
                    break;
            }

            string help = _ti.GetHelpDocumentationById(_fd.memid, out int context);
            if (0 != (_fd.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED))
            {
                lprops.Add("restricted");
            }


            if (0 != (_fd.wFuncFlags & FUNCFLAGS.FUNCFLAG_FHIDDEN))
            {
                lprops.Add("hidden");
            }

            AddHelpStringAndContext(lprops, help, context);

            if (lprops.Count > 0)
            {
                ih.AppendLine("[" + string.Join(", ", lprops.ToArray()) + "] ");
            }

            // Prototype in a different line.
            ElemDesc ed = _fd.elemdescFunc;
            GenPt paramtextgen = null;
            ElemDesc elast = null;
            bool bRetvalPresent = false;
            if (_fd.elemdescParams.Length > 0)
            {
                var names = _fd.GetNames(_ti);
                var edps = _fd.elemdescParams;
                if (edps.Length > 0)
                {
                    elast = edps[edps.Length - 1];
                }

                if (bAsDispatch && elast != null && 0 != (elast.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FRETVAL))
                {
                    bRetvalPresent = true;
                }

                int maxCnt = (bAsDispatch && bRetvalPresent) ? _fd.elemdescParams.Length - 1 : _fd.elemdescParams.Length;

                paramtextgen = x =>
                    {
                        string paramname = (names[x + 1] == null) ? "rhs" : names[x + 1];
                        ElemDesc edp = edps[x];
                        ParamDesc fd = edp.paramdesc;
                        ih.AddString(ParamFlagsDescription(edp.paramdesc) + " ");
                        edp.tdesc.ComTypeNameAsString(_ti, ih);
                        ih.AddString(" " + paramname);
                    };
            }

            (bRetvalPresent ? elast : ed).tdesc.ComTypeNameAsString(_ti, ih);
            if (memIdInSpecialRange)
            {
                ih.AddString(" " + _fd.callconv.ToString().Substring(2).ToLower());
            }

            ih.AddString(" " + _name);
            switch (_fd.elemdescParams.Length)
            {
                case 0:
                    ih.AppendLine("();");
                    break;

                case 1:
                    ih.AddString("(");
                    paramtextgen(0);
                    ih.AppendLine(");");
                    break;

                default:
                    ih.AppendLine("(");
                    using (new IDLHelperTab(ih))
                    {
                        for (int y = 0; y < _fd.elemdescParams.Length; ++y)
                        {
                            paramtextgen(y);
                            ih.AppendLine(y == _fd.elemdescParams.Length - 1 ? "" : ",");
                        }
                    }
                    ih.AppendLine(");");
                    break;
            }
        }
    }
}

