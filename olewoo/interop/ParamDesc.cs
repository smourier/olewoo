using System.Runtime.InteropServices;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class ParamDesc
    {
        private System.Runtime.InteropServices.ComTypes.PARAMDESC _desc;

        public ParamDesc(System.Runtime.InteropServices.ComTypes.PARAMDESC desc)
        {
            _desc = desc;
        }

        public System.Runtime.InteropServices.ComTypes.PARAMFLAG wParamFlags => _desc.wParamFlags;

        public object varDefaultValue
        {
            get
            {
                var ex = PtrToStructure<PARAMDESCEX>(_desc.lpVarValue);
                return ex.varDefaultValue;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PARAMDESCEX
        {
            public int cBytes;
            public object varDefaultValue;
        }
    }
}
