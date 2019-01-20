using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class FuncDesc
    {
        public FuncDesc(ITypeInfo ti, int idx)
        {
            ti.GetFuncDesc(idx, out var ptr);
            var desc = PtrToStructure<FUNCDESC>(ptr);

            if (desc.cParamsOpt != 0)
                throw new NotSupportedException();

            memid = desc.memid;
            invkind = desc.invkind;
            callconv = desc.callconv;
            wFuncFlags = (FUNCFLAGS)desc.wFuncFlags;
            elemdescFunc = new ElemDesc(desc.elemdescFunc);

            var paramsList = new List<ElemDesc>();
            for (int i = 0; i < desc.cParams; i++)
            {
                var sed = PtrToStructure<ELEMDESC>(desc.lprgelemdescParam + i * SizeOf<ELEMDESC>());
                var ed = new ElemDesc(sed);
                paramsList.Add(ed);
            }

            elemdescParams = paramsList.ToArray();
            ti.ReleaseFuncDesc(ptr);
        }

        public int memid { get; }
        public INVOKEKIND invkind { get; }
        public CALLCONV callconv { get; }
        public FUNCFLAGS wFuncFlags { get; }
        public ElemDesc elemdescFunc { get; }
        public ElemDesc[] elemdescParams { get; }
    }
}
