using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class FuncDesc
    {
        public FuncDesc(ITypeInfo ti, int idx)
        {
            ti.GetFuncDesc(idx, out var ptr);
            try
            {
                var desc = System.Runtime.InteropServices.Marshal.PtrToStructure<FUNCDESC>(ptr);
                memid = desc.memid;
                invkind = desc.invkind;
                callconv = desc.callconv;
                wFuncFlags = (FUNCFLAGS)desc.wFuncFlags;
                elemdescFunc = new ElemDesc(desc.elemdescFunc);

                var paramsList = new List<ElemDesc>();
                for (var i = 0; i < desc.cParams; i++)
                {
                    var sed = System.Runtime.InteropServices.Marshal.PtrToStructure<ELEMDESC>(desc.lprgelemdescParam + i * System.Runtime.InteropServices.Marshal.SizeOf<ELEMDESC>());
                    var ed = new ElemDesc(sed);
                    paramsList.Add(ed);
                }

                elemdescParams = paramsList.ToArray();
            }
            finally
            {
                ti.ReleaseFuncDesc(ptr);
            }
        }

        public int memid { get; }
        public INVOKEKIND invkind { get; }
        public CALLCONV callconv { get; }
        public FUNCFLAGS wFuncFlags { get; }
        public ElemDesc elemdescFunc { get; }
        public ElemDesc[] elemdescParams { get; }
    }
}
