/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Collections.Generic;
using System.Linq;

namespace olewoo
{
    class OWIDispatchMethods : TlibNode
    {
        OWDispInterface _parent;

        public OWIDispatchMethods(OWDispInterface parent)
        {
            _parent = parent;
        }

        public override string Name => "Methods";
        public override string ShortName => null;
        public override string ObjectName => null;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override int ImageIndex => (int)ImageIndices.idx_methodlist;
        public override TlibNode Parent => _parent as TlibNode;

        public override List<TlibNode> GenChildren() => _parent.MethodChildren();
        public override void BuildIDLInto(IDLFormatter ih)
        {
            var meths = Children.ToList().ConvertAll(x => x as OWMethod);
            ih.AppendLine("methods:");

            if (meths.Count > 0)
            {
                using (new IDLHelperTab(ih))
                {
                    meths.ForEach(x => x.BuildIDLInto(ih, true));
                }
            }
        }
    }
}

