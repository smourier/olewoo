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
    class OWRecord : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        ITypeInfo _ti;
        TypeAttr _ta;

        public OWRecord(TlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ti = ti;
            _ta = ta;
            _name = _ti.GetName();
        }

        public override string Name => "typedef struct " + _name;
        public override string ShortName => _name;
        public override string ObjectName => _name + "#s";
        public override int ImageIndex => (int)ImageIndices.idx_struct;
        public override TlibNode Parent => _parent;
        public override string ToString() => Name;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                var vd = new VarDesc(_ti, x);
                res.Add(new OWRecordMember(this, _ti, vd));
            }
            return res;
        }

        public override void BuildIDLInto(IDLFormatter ih)
        {
            ih.AppendLine("typedef struct tag" + _name + " {");
            using (new IDLHelperTab(ih))
            {
                Children.ForEach(x => x.BuildIDLInto(ih));
            }
            ih.AppendLine("} " + _name + ";");
        }
    }
}

