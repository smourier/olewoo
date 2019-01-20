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
    class OWDispProperty : ITlibNode
    {
        readonly ITlibNode _parent;
        readonly string _name;
        VarDesc _vd;
        ITypeInfo _ti;
        public OWDispProperty(ITlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _vd = vd;
            _ti = ti;
        }

        public override string Name => _name;                // fixme - look at varkind.
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override int ImageIndex => (int)ImageIndices.idx_strucmem;
        public override ITlibNode Parent => _parent;

        string NegStr(int x) => (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();

        public override void BuildIDLInto(IDLFormatter ih) => BuildIDLInto(ih, false);
        public override List<ITlibNode> GenChildren() => new List<ITlibNode>();
        public void BuildIDLInto(IDLFormatter ih, bool embedded)
        {
            bool memIdInSpecialRange = (_vd.memid >= 0x60000000 && _vd.memid < 0x60020000);
            var lprops = new List<string>();
            if (!memIdInSpecialRange)
            {
                lprops.Add("id(" + _vd.memid.PaddedHex() + ")");
            }

            string help = _ti.GetHelpDocumentationById(_vd.memid, out int context);
            //            if (0 != (_vd.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED)) lprops.Add("restricted");
            //            if (0 != (_vd.wFuncFlags & FUNCFLAGS.FUNCFLAG_FHIDDEN)) lprops.Add("hidden");
            AddHelpStringAndContext(lprops, help, context);
            ih.AppendLine("[" + string.Join(", ", lprops.ToArray()) + "] ");
            // Prototype in a different line.

            ElemDesc ed = _vd.elemDescVar;
            ed.tdesc.ComTypeNameAsString(_ti, ih);
            //            if (memIdInSpecialRange)
            //            {
            //                ih.AddString(" " + _fd.callconv.ToString().Substring(2).ToLower());
            //            }
            ih.AppendLine(" " + _name + ";");
        }
    }
}

