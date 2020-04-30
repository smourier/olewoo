using System;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class TypeAttr
    {
        public TypeAttr(ITypeInfo ti)
        {
            ti.GetTypeAttr(out var ptr);
            try
            {
                var attr = System.Runtime.InteropServices.Marshal.PtrToStructure<TYPEATTR>(ptr);
                typekind = attr.typekind;
                wTypeFlags = attr.wTypeFlags;
                cImplTypes = attr.cImplTypes;
                cFuncs = attr.cFuncs;
                cVars = attr.cVars;
                tdescAlias = new TypeDesc(attr.tdescAlias);
                guid = attr.guid;
                wMajorVerNum = attr.wMajorVerNum;
                wMinorVerNum = attr.wMinorVerNum;
            }
            finally
            {
                ti.ReleaseTypeAttr(ptr);
            }
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
            var ptr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(IntPtr.Size);
            ti.GetDllEntry(memid, invKind, ptr, IntPtr.Zero, IntPtr.Zero);
            var bstr = System.Runtime.InteropServices.Marshal.ReadIntPtr(ptr);
            var name = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(bstr);
            System.Runtime.InteropServices.Marshal.FreeBSTR(bstr);
            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(ptr);
            return name;
        }
    }
}
