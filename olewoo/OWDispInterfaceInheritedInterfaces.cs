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
    // A dispinterface's first inherited interface is the swap for interface.
    class OWDispInterfaceInheritedInterfaces : OWInheritedInterfaces
    {
        public OWDispInterfaceInheritedInterfaces(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
            : base(parent, ti, ta)
        {
        }

        public override List<ITlibNode> GenChildren()
        {
            ITypeInfo ti = _ti;
            TypeAttr ta = _ta;
            ITypeInfoXtra.SwapForInterface(ref ti, ref ta);
            var res = new List<ITlibNode>();
            res.Add(new OWInterface(this, ti, ta, false));
            return res;
        }
    }
}

