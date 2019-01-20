using System;
using System.Runtime.InteropServices;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class ParamDesc
    {
        public ParamDesc(System.Runtime.InteropServices.ComTypes.PARAMDESC desc)
        {
            wParamFlags = desc.wParamFlags;
            if (desc.lpVarValue != IntPtr.Zero)
            {
                var ex = PtrToStructure<PARAMDESCEX>(desc.lpVarValue);
                varDefaultValue = ex.varDefaultValue;
            }
        }

        public System.Runtime.InteropServices.ComTypes.PARAMFLAG wParamFlags { get; }
        public object varDefaultValue { get; }

        [StructLayout(LayoutKind.Sequential)]
        private struct PARAMDESCEX
        {
            public int cBytes;
            public object varDefaultValue;
        }
    }
}
