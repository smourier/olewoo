﻿/**************************************
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
    class OWIDispatchProperties : TlibNode
    {
        OWDispInterface _parent;

        public OWIDispatchProperties(OWDispInterface parent)
        {
            _parent = parent;
        }

        public override string Name => "Properties";
        public override string ShortName => null;
        public override string ObjectName => null;
        public override int ImageIndex => (int)ImageIndices.idx_propertylist;
        public override TlibNode Parent => _parent as TlibNode;
        public override string ToString() => Name;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override List<TlibNode> GenChildren() => _parent.PropertyChildren();
        public override void BuildIDLInto(IDLFormatter ih)
        {
            var props = Children.ToList().ConvertAll(x => x as OWDispProperty);
            ih.AppendLine("properties:");

            if (props.Count > 0)
            {
                using (new IDLHelperTab(ih))
                {
                    props.ForEach(x => x.BuildIDLInto(ih, true));
                }
            }
        }
    }
}

