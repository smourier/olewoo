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
    class OWEnumValue : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        VarDesc _vd;
        ITypeInfo _ti;
        readonly int _val;

        public OWEnumValue(TlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _val = (int)vd.varValue;
            _vd = vd;
            _ti = ti;
        }

        public override string Name => "const int " + _name + " = " + _val; // fixme - look at varkind.
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override int ImageIndex => (int)ImageIndices.idx_const;
        public override TlibNode Parent => _parent;

        string NegStr(int x) => (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();

        public override void BuildIDLInto(IDLFormatter ih) => BuildIDLInto(ih, false, false);
        public override List<TlibNode> GenChildren() => new List<TlibNode>();
        public void BuildIDLInto(IDLFormatter ih, bool embedded, bool islast) => ih.AppendLine("const int " + _ti.GetDocumentationById(_vd.memid) + " = " + NegStr(_val) + (embedded ? (islast ? "" : ",") : ";"));
    }
}

