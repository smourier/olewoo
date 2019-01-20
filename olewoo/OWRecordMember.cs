﻿/**************************************
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
    class OWRecordMember : ITlibNode
    {
        readonly ITlibNode _parent;
        readonly string _type;
        readonly string _name;
        VarDesc _vd;
        readonly ITypeInfo _ti;

        public OWRecordMember(ITlibNode parent, ITypeInfo ti, VarDesc vd)
        {
            _parent = parent;
            _name = ti.GetDocumentationById(vd.memid);
            _vd = vd;
            _ti = ti;
            var ig = new IDLGrabber();
            _vd.elemDescVar.tdesc.ComTypeNameAsString(_ti, ig);
            _type = ig.Value;
        }

        public override string Name => _type + " " + _name;
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override int ImageIndex => (int)ImageIndices.idx_strucmem;
        public override ITlibNode Parent => _parent;

        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => false;
        public override List<ITlibNode> GenChildren() => new List<ITlibNode>();
        public override void BuildIDLInto(IDLFormatter ih) => ih.AppendLine(Name + ";");
    }
}

