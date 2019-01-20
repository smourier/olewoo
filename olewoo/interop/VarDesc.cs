using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class VarDesc
    {
        private VARDESC _desc;

        public VarDesc(ITypeInfo ti, int idx)
        {
            ti.GetVarDesc(idx, out var ptr);
            _desc = PtrToStructure<VARDESC>(ptr);
            ti.ReleaseVarDesc(ptr);
        }

        public int memid => _desc.memid;
        public ElemDesc elemDescVar => new ElemDesc(_desc.elemdescVar);
        public object varValue => GetObjectForNativeVariant(_desc.desc.lpvarValue);
    }
}
