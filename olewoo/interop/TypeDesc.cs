using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo.interop
{
    public class TypeDesc
    {
        private System.Runtime.InteropServices.ComTypes.TYPEDESC _desc;

        public TypeDesc(System.Runtime.InteropServices.ComTypes.TYPEDESC desc)
        {
            _desc = desc;
        }

        public int hreftype => (int)(ulong)_desc.lpValue;

        public void ComTypeNameAsString(ITypeInfo ti, IDLFormatter_iop ift)
        {
            var vr = (VarEnum)_desc.vt;
            System.Runtime.InteropServices.ComTypes.TYPEDESC pd;
            switch (vr)
            {
                case VarEnum.VT_PTR:
                    pd = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEDESC>(_desc.lpValue);
                    new TypeDesc(pd).ComTypeNameAsString(ti, ift);
                    ift.AddString("*");
                    break;

                case VarEnum.VT_SAFEARRAY:
                    ift.AddString("SAFEARRAY(");
                    pd = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEDESC>(_desc.lpValue);
                    new TypeDesc(pd).ComTypeNameAsString(ti, ift);
                    ift.AddString(")");
                    break;

                case VarEnum.VT_USERDEFINED:
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
                    break;

                case VarEnum.VT_I2:
                    ift.AddString("short");
                    break;

                case VarEnum.VT_I4:
                    ift.AddString("long");
                    break;

                case VarEnum.VT_R4:
                    ift.AddString("float");
                    break;

                case VarEnum.VT_R8:
                    ift.AddString("double");
                    break;

                case VarEnum.VT_CY:
                    ift.AddString("CY");
                    break;

                case VarEnum.VT_DATE:
                    ift.AddString("DATE");
                    break;

                case VarEnum.VT_BSTR:
                    ift.AddString("BSTR");
                    break;

                case VarEnum.VT_DISPATCH:
                    ift.AddString("IDispatch*");
                    break;

                case VarEnum.VT_ERROR:
                    ift.AddString("SCODE");
                    break;

                case VarEnum.VT_BOOL:
                    ift.AddString("VARIANT_BOOL");
                    break;

                case VarEnum.VT_VARIANT:
                    ift.AddString("VARIANT");
                    break;

                case VarEnum.VT_UNKNOWN:
                    ift.AddString("IUnknown*");
                    break;

                case VarEnum.VT_UI1:
                    ift.AddString("BYTE");
                    break;

                case VarEnum.VT_DECIMAL:
                    ift.AddString("DECIMAL");
                    break;

                case VarEnum.VT_I1:
                    ift.AddString("char");
                    break;

                case VarEnum.VT_UI2:
                    ift.AddString("unsigned short");
                    break;

                case VarEnum.VT_UI4:
                    ift.AddString("unsigned long");
                    break;

                case VarEnum.VT_I8:
                    ift.AddString("__int64");
                    break;

                case VarEnum.VT_UI8:
                    ift.AddString("unsigned __int64");
                    break;

                case VarEnum.VT_INT:
                    ift.AddString("int");
                    break;

                case VarEnum.VT_UINT:
                    ift.AddString("unsigned int");
                    break;

                case VarEnum.VT_HRESULT:
                    ift.AddString("HRESULT");
                    break;

                case VarEnum.VT_VOID:
                    ift.AddString("void");
                    break;

                case VarEnum.VT_LPSTR:
                    ift.AddString("LPSTR");
                    break;

                case VarEnum.VT_LPWSTR:
                    ift.AddString("LPWSTR");
                    break;

                case VarEnum.VT_FILETIME:
                    ift.AddString("FILETIME");
                    break;

                case VarEnum.VT_STREAM:
                    ift.AddString("IStream*");
                    break;

                case VarEnum.VT_STORAGE:
                    ift.AddString("IStorage*");
                    break;

                case VarEnum.VT_CLSID:
                    ift.AddString("CLSID");
                    break;

                default:
                    ift.AddString("!!!" + vr + "!!!");
                    break;
            }
        }
    }
}
