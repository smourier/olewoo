using System;
using System.Runtime.InteropServices;

namespace olewoo.interop
{
    public class ParamDesc
    {
        public ParamDesc(System.Runtime.InteropServices.ComTypes.PARAMDESC desc)
        {
            wParamFlags = desc.wParamFlags;
            if (desc.wParamFlags.HasFlag(System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FHASDEFAULT) ||
                desc.wParamFlags.HasFlag(System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FOPT))
            {
                if (desc.lpVarValue != IntPtr.Zero)
                {
                    //typedef struct tagPARAMDESCEX
                    //{
                    //    ULONG cBytes;
                    //    VARIANTARG varDefaultValue;
                    //}
                    //PARAMDESCEX, *LPPARAMDESCEX;

                    // for some reason, padding is set to 8 even in x86
                    varDefaultValue = Marshal.GetObjectForNativeVariant(desc.lpVarValue + 8);
                }
            }
        }

        public System.Runtime.InteropServices.ComTypes.PARAMFLAG wParamFlags { get; }
        public object varDefaultValue { get; }

   }
}
