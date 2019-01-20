using System;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class TypeLibAttr
    {
        private TYPELIBATTR _attr;

        public TypeLibAttr(ITypeLib tl)
        {
            tl.GetLibAttr(out var ptr);
            _attr = PtrToStructure<TYPELIBATTR>(ptr);
        }

        public Guid guid => _attr.guid;
        public short wMajorVerNum => _attr.wMajorVerNum;
        public short wMinorVerNum => _attr.wMinorVerNum;
    }
}
