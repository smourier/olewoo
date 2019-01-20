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
    class OWModule : ITlibNode
    {
        readonly ITlibNode _parent;
        readonly string _name;
        ITypeInfo _ti;
        TypeAttr _ta;
        readonly string _dllname;

        public OWModule(ITlibNode parent, ITypeInfo ti, TypeAttr ta)
        {
            _parent = parent;
            _ti = ti;
            _ta = ta;
            _name = _ti.GetName();
            if (_ta.cVars > 0 || _ta.cFuncs > 0)
            {
                int memid;
                var invkind = INVOKEKIND.INVOKE_FUNC;
                if (_ta.cFuncs > 0)
                {
                    var fd = new FuncDesc(_ti, 0);
                    invkind = fd.invkind;
                    memid = fd.memid;
                    _dllname = _ta.GetDllEntry(_ti, invkind, memid);
                }
                else
                {
                    _dllname = null;
                }
            }
        }

        public override string Name => "module " + _name;
        public override string ShortName => _name;
        public override string ObjectName => null;
        public override bool DisplayAtTLBLevel(ICollection<string> interfaceNames) => true;
        public override int ImageIndex => (int)ImageIndices.idx_module;
        public override ITlibNode Parent => _parent;

        public override List<ITlibNode> GenChildren()
        {
            var res = new List<ITlibNode>();
            if (_ta.cVars > 0)
            {
                res.Add(new OWChildrenIndirect(this, "Constants", (int)ImageIndices.idx_constlist, GenConstChildren));
            }

            if (_ta.cFuncs > 0)
            {
                res.Add(new OWChildrenIndirect(this, "Functions", (int)ImageIndices.idx_methodlist, GenFuncChildren));
            }
            return res;
        }

        private List<ITlibNode> GenConstChildren()
        {
            var res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cVars; ++x)
            {
                var vd = new VarDesc(_ti, x);
                res.Add(new OWModuleConst(this, _ti, vd, x));
            }
            return res;
        }

        private List<ITlibNode> GenFuncChildren()
        {
            var res = new List<ITlibNode>();
            for (int x = 0; x < _ta.cFuncs; ++x)
            {
                var fd = new FuncDesc(_ti, x);
                res.Add(new OWMethod(this, _ti, fd));
            }
            return res;
        }

        public override void BuildIDLInto(IDLFormatter ih)
        {
            if (_ta.cFuncs == 0)
            {
                ih.AppendLine("// NOTE: This module has no entry points. There is no way to");
                ih.AppendLine("//       extract the dllname of a module with no entry points!");
                ih.AppendLine("// ");
            }

            ih.AppendLine("[");
            var liba = new List<string>();
            liba.Add("dllname(\"" + (string.IsNullOrEmpty(_dllname) ? "<no entry points>" : _dllname) + "\")");

            if (_ta.guid != Guid.Empty)
            {
                liba.Add("uuid(" + _ta.guid + ")");
            }

            string help = _ti.GetHelpDocumentationById(-1, out int cnt);
            if (!string.IsNullOrEmpty(help))
            {
                liba.Add("helpstring(\"" + help + "\")");
            }

            if (cnt != 0)
            {
                liba.Add("helpcontext(" + cnt.PaddedHex() + ")");
            }

            cnt = 0;
            liba.ForEach(x => ih.AppendLine("  " + x + (++cnt == liba.Count ? "" : ",")));
            ih.AppendLine("]");
            ih.AppendLine("module " + _name + " {");
            using (new IDLHelperTab(ih))
            {
                Children.ForEach(x => x.BuildIDLInto(ih));
            }
            ih.AppendLine("};");
        }
    }
}

