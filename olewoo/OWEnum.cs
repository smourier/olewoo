/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using olewoo.interop;

namespace olewoo
{
    class OWEnum : TlibNode
    {
        readonly TlibNode _parent;
        readonly string _name;
        TypeAttr _ta;
        readonly ITypeInfo _ti;

        public OWEnum(TlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _name = ti.GetName();
            _ta = ta;
            _ti = ti;
        }

        public override string Name => "typedef enum " + _name;
        public override string ShortName => _name;
        public override string ObjectName => _name + "#i";
        public override int ImageIndex => (int)ImageIndices.idx_enum;
        public override TlibNode Parent => _parent;
        public override string ToString() => Name;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                var vd = new VarDesc(_ti, x);
                res.Add(new OWEnumValue(this, _ti, vd));
            }
            return res;
        }

        public override void BuildIDLInto(IDLFormatter ih)
        {
            string tde = "typedef ";
            // If the enum has a uuid, or a version associate with it, we provide that information on the same line.

            if (!_ta.guid.Equals(Guid.Empty))
            {
                tde += "[uuid(" + _ta.guid + "), version(" + _ta.wMajorVerNum + "." + _ta.wMinorVerNum + ")]";
                ih.AppendLine(tde);
                tde = "";
            }

            ih.AppendLine(tde + "enum {");
            using (new IDLHelperTab(ih))
            {
                int idx = 0;
                Children.ForEach(x => (x as OWEnumValue).BuildIDLInto(ih, true, ++idx == _ta.cVars));
            }
            ih.AppendLine("} " + _name + ";");
        }
    }
}

