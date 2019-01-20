using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class CustomDatas
    {
        private readonly List<CUSTDATAITEM> _items = new List<CUSTDATAITEM>();

        public CustomDatas(ITypeLib2 tl)
        {
            var ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf<CUSTDATA>());
            tl.GetAllCustData(ptr);
            var cd = Marshal.PtrToStructure<CUSTDATA>(ptr);
            for (int i = 0; i < cd.cCustData; i++)
            {
            }
            ClearCustData(ptr);
        }

        public IReadOnlyList<CUSTDATAITEM> Items => _items;

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
