using System;
using System.Runtime.InteropServices;

namespace olewoo.interop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct CUSTDATAITEM
    {
        public Guid guid;
        public object varValue;
    }
}
