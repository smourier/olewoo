using System;
using System.Runtime.InteropServices.ComTypes;
using static System.Runtime.InteropServices.Marshal;

namespace olewoo.interop
{
    public class TypeAttr
    {
        public TypeAttr(ITypeInfo ti)
        {
            ti.GetTypeAttr(out var ptr);
            var attr = PtrToStructure<TYPEATTR>(ptr);
            typekind = attr.typekind;
            wTypeFlags = attr.wTypeFlags;
            cImplTypes = attr.cImplTypes;
            cFuncs = attr.cFuncs;
            cVars = attr.cVars;
            tdescAlias = new TypeDesc(attr.tdescAlias);
            guid = attr.guid;
            wMajorVerNum = attr.wMajorVerNum;
            wMinorVerNum = attr.wMinorVerNum;
            ti.ReleaseTypeAttr(ptr);
        }

        public TYPEKIND typekind { get; }
        public TYPEFLAGS wTypeFlags { get; }
        public int cImplTypes { get; }
        public int cFuncs { get; }
        public int cVars { get; }
        public TypeDesc tdescAlias { get; }
        public Guid guid { get; }
        public short wMajorVerNum { get; }
        public short wMinorVerNum { get; }

        public string GetDllEntry(ITypeInfo ti, INVOKEKIND invKind, int memid)
        {
            var ptr = AllocCoTaskMem(IntPtr.Size);
            ti.GetDllEntry(memid, invKind, ptr, IntPtr.Zero, IntPtr.Zero);
            var bstr = ReadIntPtr(ptr);
            var name = PtrToStringBSTR(bstr);
            FreeBSTR(bstr);
            FreeCoTaskMem(ptr);
            return name;
        }
    }
}
