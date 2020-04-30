using System;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class TypeLibAttr
    {
        public TypeLibAttr(ITypeLib tl)
        {
            tl.GetLibAttr(out var ptr);
            try
            {
                var attr = PtrToStructure<TYPELIBATTR>(ptr);
                guid = attr.guid;
                wMajorVerNum = attr.wMajorVerNum;
                wMinorVerNum = attr.wMinorVerNum;
                lcid = attr.lcid;
                syskind = attr.syskind;
                wLibFlags = attr.wLibFlags;
            }
            finally
            {
                tl.ReleaseTLibAttr(ptr);
            }
        }

        public Guid guid { get; }
        public int lcid { get; }
        public SYSKIND syskind { get; }
        public short wMajorVerNum { get; }
        public short wMinorVerNum { get; }
        public LIBFLAGS wLibFlags { get; }
    }
}
