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
            var items = new List<CUSTDATAITEM>();
            if (tl != null)
            {
                IntPtr ptr;
                if (IntPtr.Size == 8)
                {
                    ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf<CUSTDATA_64>());
                }
                else
                {
                    ptr = Marshal.AllocCoTaskMem(Marshal.SizeOf<CUSTDATA_32>());
                }

                try
                {
                    tl.GetAllCustData(ptr);
                    try
                    {
                        int count;
                        IntPtr prgCustData;
                        if (IntPtr.Size == 8)
                        {
                            var cd64 = Marshal.PtrToStructure<CUSTDATA_64>(ptr);
                            count = cd64.cCustData;
                            prgCustData = cd64.prgCustData;
                        }
                        else
                        {
                            var cd32 = Marshal.PtrToStructure<CUSTDATA_32>(ptr);
                            count = cd32.cCustData;
                            prgCustData = cd32.prgCustData;
                        }

                        for (int i = 0; i < count; i++)
                        {
                            var p = prgCustData + Marshal.SizeOf<CUSTDATAITEM>() * i;
                            var guid = Marshal.PtrToStructure<Guid>(p);
                            var o = Marshal.GetObjectForNativeVariant(p + Marshal.SizeOf<Guid>());
                            items.Add(new CUSTDATAITEM { guid = guid, varValue = o });
                        }
                    }
                    finally
                    {
                        ClearCustData(ptr);
                    }
                }
                finally
                {
                    Marshal.FreeCoTaskMem(ptr);
                }
            }
            Items = items.ToArray();
        }

        public CUSTDATAITEM[] Items { get; }

        [DllImport("oleaut32")]
        private static extern void ClearCustData(IntPtr ptr);

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct CUSTDATA_32
        {
            public int cCustData;
            public IntPtr prgCustData;
        }

        // for some reason, we must set pack = 8 here
        [StructLayout(LayoutKind.Sequential, Pack = 8)]
        private struct CUSTDATA_64
        {
            public int cCustData;
            public IntPtr prgCustData;
        }
    }
}
