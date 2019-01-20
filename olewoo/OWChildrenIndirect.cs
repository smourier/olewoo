/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Collections.Generic;

namespace olewoo
{
    class OWChildrenIndirect : TlibNode
    {
        readonly TlibNode _parent;
        readonly int _imageidx;
        readonly string _name;
        readonly dlgCreateChildren _genChildren;

        public OWChildrenIndirect(TlibNode parent, string name, int imageidx, dlgCreateChildren genchildren)
        {
            _parent = parent;
            _name = name;
            _imageidx = imageidx;
            _genChildren = genchildren;
        }

        public override string Name => _name;
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override int ImageIndex => _imageidx;
        public override TlibNode Parent => _parent;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override List<TlibNode> GenChildren() => _genChildren();
        public override void BuildIDLInto(IDLFormatter ih) => Children.ForEach(x => x.BuildIDLInto(ih));
    }
}

