using System;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class TypeAttr
    {
        private TYPEATTR _attr;

        public TypeAttr(ITypeInfo ti)
        {
            ti.GetTypeAttr(out var ptr);
            _attr = PtrToStructure<TYPEATTR>(ptr);
            ti.ReleaseTypeAttr(ptr);
        }

        public TYPEKIND typekind => _attr.typekind;
        public TYPEFLAGS wTypeFlags => _attr.wTypeFlags;
        public int cImplTypes => _attr.cImplTypes;
        public int cFuncs => _attr.cFuncs;
        public int cVars => _attr.cVars;
        public TypeDesc tdescAlias => new TypeDesc(_attr.tdescAlias);
        public Guid guid => _attr.guid;
        public short wMajorVerNum => _attr.wMajorVerNum;
        public short wMinorVerNum => _attr.wMinorVerNum;

        public string GetDllEntry(ITypeInfo ti, INVOKEKIND invKind, int memid)
        {
            var bstr = AllocCoTaskMem(IntPtr.Size);
            ti.GetDllEntry(memid, invKind, bstr, IntPtr.Zero, IntPtr.Zero);
            var name = PtrToStringBSTR(bstr);
            FreeBSTR(bstr);
            return name;
        }
    }
}
