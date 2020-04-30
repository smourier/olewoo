using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class VarDesc
    {
        public VarDesc(ITypeInfo ti, int idx)
        {
            ti.GetVarDesc(idx, out var ptr);
            try
            {
                var desc = System.Runtime.InteropServices.Marshal.PtrToStructure<VARDESC>(ptr);
                memid = desc.memid;
                elemDescVar = new ElemDesc(desc.elemdescVar);
                if (desc.varkind == VARKIND.VAR_CONST)
                {
                    varValue = System.Runtime.InteropServices.Marshal.GetObjectForNativeVariant(desc.desc.lpvarValue);
                }
            }
            finally
            {
                ti.ReleaseVarDesc(ptr);
            }
        }

        public int memid { get; }
        public ElemDesc elemDescVar { get; }
        public object varValue { get; }
    }
}
