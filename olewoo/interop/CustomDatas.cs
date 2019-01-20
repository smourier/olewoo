using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class CustomDatas
    {
        public CustomDatas(ITypeLib2 tl)
        {
            var ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf<CUSTDATA>());
            tl.GetAllCustData(ptr);
            var cd = Marshal.PtrToStructure<CUSTDATA>(ptr);
            var items = new List<CUSTDATAITEM>();
            for (int i = 0; i < cd.cCustData; i++)
            {
            }
            Items = items.ToArray();
            ClearCustData(ptr);
        }

        public CUSTDATAITEM[] Items { get; }

        [DllImport("oleaut32")]
        private static extern void ClearCustData(IntPtr ptr);

        [StructLayout(LayoutKind.Sequential)]
        private struct CUSTDATA
        {
            public int cCustData;
            public IntPtr prgCustData;
        }
    }
}
