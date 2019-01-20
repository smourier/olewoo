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
    class OWModuleConst : ITlibNode
    {
        readonly ITlibNode _parent;
        readonly string _name;
        VarDesc _vd;
        readonly ITypeInfo _ti;
        object _val;
        readonly int _idx;

        public OWModuleConst(ITlibNode parent, ITypeInfo ti, VarDesc vd, int idx)
        {
            _parent = parent;
            _vd = vd;
            _ti = ti;
            var ig = new IDLGrabber();
            _vd.elemDescVar.tdesc.ComTypeNameAsString(_ti, ig);
            _name = ig.Value + " " + ti.GetDocumentationById(vd.memid);
            _val = vd.varValue;
            if (_val == null)
            {
                _val = "";
            }

            if (_val.GetType() == typeof(string))
            {
                _val = (_val as string).ReEscape();
            }

            _idx = idx;
        }

        public override string Name => "const " + _name + " = " + _val;
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override int ImageIndex => (int)ImageIndices.idx_const;
        public override ITlibNode Parent => _parent;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override List<ITlibNode> GenChildren() => new List<ITlibNode>();

        string NegStr(int x) => (x < 0) ? ("0x" + x.ToString("X")) : x.ToString();

        public override void BuildIDLInto(IDLFormatter ih) => BuildIDLInto(ih, false, false);
        public void BuildIDLInto(IDLFormatter ih, bool embedded, bool islast)
        {
            string desc = "";
            //int cnt = 0;
            //String help = _ti.GetHelpDocumentationById(_idx, out cnt);
            //List<String> props = new List<string>();
            //AddHelpStringAndContext(props, help, cnt);
            //if (props.Count > 0)
            //{
            //    desc += "[" + String.Join(",", props.ToArray()) + "] ";
            //}
            desc += _val.GetType() == typeof(int) ? NegStr((int)_val) : _val.ToString();
            ih.AppendLine("const " + _name + " = " + desc + (embedded ? (islast ? "" : ",") : ";"));
        }
    }
}

