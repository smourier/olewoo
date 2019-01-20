using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class VarDesc
    {
        public VarDesc(ITypeInfo ti, int idx)
        {
            ti.GetVarDesc(idx, out var ptr);
            var desc = PtrToStructure<VARDESC>(ptr);
            memid = desc.memid;
            elemDescVar = new ElemDesc(desc.elemdescVar);
            if (desc.varkind == VARKIND.VAR_CONST)
            {
                varValue = GetObjectForNativeVariant(desc.desc.lpvarValue);
            }
            ti.ReleaseVarDesc(ptr);
        }

        public int memid { get; }
        public ElemDesc elemDescVar { get; }
        public object varValue { get; }
    }
}
