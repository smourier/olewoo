using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class FuncDesc
    {
        private readonly FUNCDESC _desc;

        public FuncDesc(ITypeInfo ti, int idx)
        {
            ti.GetFuncDesc(idx, out var ptr);
            _desc = PtrToStructure<FUNCDESC>(ptr);

            if (_desc.cParamsOpt != 0)
                throw new NotSupportedException();

            var paramsList = new List<ElemDesc>();
            for (int i = 0; i < _desc.cParams; i++)
            {
                var sed = PtrToStructure<ELEMDESC>(_desc.lprgelemdescParam + i * SizeOf<ELEMDESC>());
                var ed = new ElemDesc(sed);
                paramsList.Add(ed);
            }

            elemdescParams = paramsList.ToArray();
            ti.ReleaseFuncDesc(ptr);
        }

        public int memid => _desc.memid;
        public INVOKEKIND invkind => _desc.invkind;
        public int cParams => _desc.cParams;
        public CALLCONV callconv => _desc.callconv;
        public FUNCFLAGS wFuncFlags => (FUNCFLAGS)_desc.wFuncFlags;
        public ElemDesc elemdescFunc => new ElemDesc(_desc.elemdescFunc);

        public ElemDesc[] elemdescParams { get; }
    }
}
