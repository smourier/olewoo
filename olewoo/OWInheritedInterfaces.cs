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
    class OWInheritedInterfaces : TlibNode
    {
        readonly TlibNode _parent;
        protected TypeAttr _ta;
        protected ITypeInfo _ti;

        public OWInheritedInterfaces(TlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ta = ta;
            _ti = ti;
        }

        public override string Name => "Inherited Interfaces";
        public override string ShortName => null;
        public override string ObjectName => null;
        public override int ImageIndex => (int)ImageIndices.idx_interface;
        public override TlibNode Parent => _parent;
        public override string ToString() => Name;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override void BuildIDLInto(IDLFormatter ih) => ih.AppendLine("");
        public override List<TlibNode> GenChildren()
        {
            var res = new List<TlibNode>();

            if (_ta.cImplTypes > 0)
            {
                if (_ta.cImplTypes > 1)
                    throw new Exception("Multiple inheritance!?");

                _ti.GetRefTypeOfImplType(0, out int href);
                _ti.GetRefTypeInfo(href, out ITypeInfo ti2);
                CommonBuildTlibNode(this, ti2, false, true, res);
            }
            return res;
        }
    }
}

