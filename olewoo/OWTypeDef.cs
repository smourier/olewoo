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
    class OWTypeDef : TlibNode
    {
        readonly TlibNode _parent;
        ITypeInfo _ti;
        TypeAttr _ta;
        readonly string _name;

        public OWTypeDef(TlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ta = ta;
            _ti = ti;

            ITypeInfo oti = null;
            try
            {
                _ti.GetRefTypeInfo(_ta.tdescAlias.hreftype, out oti);
            }
            catch
            {
            }

            if (oti != null)
            {
                _name = oti.GetName() + " " + ti.GetName();
            }
        }

        public override string Name => "typedef " + _name;
        public override string ShortName => _name;
        public override string ObjectName => _name + "#i";
        public override int ImageIndex => (int)ImageIndices.idx_typedef;
        public override TlibNode Parent => _parent;

        public override void BuildIDLInto(IDLFormatter ih) => ih.AppendLine("typedef [public] " + _name + ";");
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();
            ITypeInfo oti = null;
            try
            {
                _ti.GetRefTypeInfo(_ta.tdescAlias.hreftype, out oti);
            }
            catch
            {
            }

            // fixed infinite recursion
            if (oti != null && _ti != oti)
            {
                CommonBuildTlibNode(this, oti, false, false, res);
            }
            return res;
        }
    }
}

