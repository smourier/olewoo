using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class TypeDesc
    {
        private readonly Action<ITypeInfo, IDLFormatter, VarEnum> _action;

        public TypeDesc(System.Runtime.InteropServices.ComTypes.TYPEDESC desc)
        {
            hreftype = (int)(ulong)desc.lpValue;
            Type = (VarEnum)desc.vt;
            System.Runtime.InteropServices.ComTypes.TYPEDESC pd;
            switch (Type)
            {
                case VarEnum.VT_PTR:
                    pd = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEDESC>(desc.lpValue);
                    _action = (ti, ift, v) =>
                    {
                        new TypeDesc(pd).ComTypeNameAsString(ti, ift);
                        ift.AddString("*");
                    };
                    break;

                case VarEnum.VT_SAFEARRAY:
                    pd = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEDESC>(desc.lpValue);
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("SAFEARRAY(");
                        new TypeDesc(pd).ComTypeNameAsString(ti, ift);
                        ift.AddString(")");
                    };
                    break;

                case VarEnum.VT_CARRAY:
                    var ad = Marshal.PtrToStructure<ARRAYDESC>(desc.lpValue);
                    _action = (ti, ift, v) =>
                    {
                        for (int i = 0; i < ad.cDims; i++)
                        {
                            var bounds = Marshal.PtrToStructure<SAFEARRAYBOUND>(ad.rgbounds + Marshal.SizeOf<SAFEARRAYBOUND>() * i);
                            ift.AddString("[");
                            ift.AddString(bounds.lLbound.ToString());
                            ift.AddString("...");
                            ift.AddString((bounds.cElements + bounds.lLbound - 1).ToString());
                            ift.AddString("]");
                        }
                    };
                    break;

                case VarEnum.VT_USERDEFINED:
                    _action = (ti, ift, v) =>
                    {
                        ITypeInfo cti = null;
                        try
                        {
                            ti.GetRefTypeInfo(hreftype, out cti);
                        }
                        catch
                        {
                        }

                        string name;
                        if (cti != null)
                        {
                            cti.GetDocumentation(-1, out name, out var docSrting, out var ctx, out var helpFile);
                        }
                        else
                        {
                            name = "???";
                        }

                        ift.AddLink(name, "i");
                    };
                    break;

                case VarEnum.VT_I2:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("short");
                    };
                    break;

                case VarEnum.VT_I4:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("long");
                    };
                    break;

                case VarEnum.VT_R4:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("float");
                    };
                    break;

                case VarEnum.VT_R8:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("double");
                    };
                    break;

                case VarEnum.VT_CY:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("CY");
                    };
                    break;

                case VarEnum.VT_DATE:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("DATE");
                    };
                    break;

                case VarEnum.VT_BSTR:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("BSTR");
                    };
                    break;

                case VarEnum.VT_DISPATCH:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("IDispatch*");
                    };
                    break;

                case VarEnum.VT_ERROR:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("SCODE");
                    };
                    break;

                case VarEnum.VT_BOOL:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("VARIANT_BOOL");
                    };
                    break;

                case VarEnum.VT_VARIANT:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("VARIANT");
                    };
                    break;

                case VarEnum.VT_UNKNOWN:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("IUnknown*");
                    };
                    break;

                case VarEnum.VT_UI1:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("BYTE");
                    };
                    break;

                case VarEnum.VT_DECIMAL:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("DECIMAL");
                    };
                    break;

                case VarEnum.VT_I1:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("char");
                    };
                    break;

                case VarEnum.VT_UI2:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("unsigned short");
                    };
                    break;

                case VarEnum.VT_UI4:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("unsigned long");
                    };
                    break;

                case VarEnum.VT_I8:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("__int64");
                    };
                    break;

                case VarEnum.VT_UI8:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("unsigned __int64");
                    };
                    break;

                case VarEnum.VT_INT:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("int");
                    };
                    break;

                case VarEnum.VT_UINT:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("unsigned int");
                    };
                    break;

                case VarEnum.VT_HRESULT:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("HRESULT");
                    };
                    break;

                case VarEnum.VT_VOID:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("void");
                    };
                    break;

                case VarEnum.VT_LPSTR:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("LPSTR");
                    };
                    break;

                case VarEnum.VT_LPWSTR:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("LPWSTR");
                    };
                    break;

                case VarEnum.VT_FILETIME:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("FILETIME");
                    };
                    break;

                case VarEnum.VT_STREAM:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("IStream*");
                    };
                    break;

                case VarEnum.VT_STORAGE:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("IStorage*");
                    };
                    break;

                case VarEnum.VT_CLSID:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("CLSID");
                    };
                    break;

                default:
                    _action = (ti, ift, v) =>
                    {
                        ift.AddString("!!! " + v + " !!!");
                    };
                    break;
            }
        }

        public int hreftype { get; }
        public VarEnum Type { get; }

        public void ComTypeNameAsString(ITypeInfo ti, IDLFormatter ift) => _action(ti, ift, Type);

        [StructLayout(LayoutKind.Sequential)]
        private struct ARRAYDESC
        {
            public System.Runtime.InteropServices.ComTypes.TYPEDESC tdescElem;
            public short cDims;
            public IntPtr rgbounds;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SAFEARRAYBOUND
        {
            public int cElements;
            public int lLbound;
        }
    }
}
