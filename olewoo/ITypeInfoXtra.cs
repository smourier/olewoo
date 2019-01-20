/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Runtime.InteropServices.ComTypes;
using olewoo.interop;

namespace olewoo
{
    static class ITypeInfoXtra
    {
        private const int MEMBERID_NONE = -1;

        public static TypeAttr GetTypeAttr(this ITypeInfo ti) => new TypeAttr(ti);
        public static string PaddedHex(this int x) => "0x" + x.ToString("x").PadLeft(8, '0');
        public static string GetName(this ITypeInfo ti) => ti.GetDocumentationById(MEMBERID_NONE);
        public static string GetName(this ITypeLib ti)
        {
            ti.GetDocumentation(MEMBERID_NONE, out string res, out string ignored, out int ctx, out string ignored2);
            return res;
        }

        public static string GetDocumentationById(this ITypeInfo ti, int memid)
        {
            ti.GetDocumentation(memid, out string res, out string ignored, out int ctx, out string ignored2);
            return res;
        }

        public static string GetHelpDocumentationById(this ITypeInfo ti, int memid, out int context)
        {
            ti.GetDocumentation(memid, out string ignored, out string res, out context, out string ignored2);
            return res;
        }

        public static string GetHelpDocumentation(this ITypeLib ti, out int context)
        {
            ti.GetDocumentation(MEMBERID_NONE, out string ignored, out string res, out context, out string ignored2);
            return res;
        }

        public static bool SwapForInterface(ref ITypeInfo ti, ref TypeAttr ta)
        {
            if (ta.typekind == TYPEKIND.TKIND_DISPATCH && 0 != (ta.wTypeFlags & TYPEFLAGS.TYPEFLAG_FDUAL))
            {
                ti.GetRefTypeOfImplType(-1, out int href);
                ti.GetRefTypeInfo(href, out ti);
                ta = new TypeAttr(ti);
                return true;
            }
            return false;
        }

        public static string QuoteString(object obj)
        {
            if (obj == null)
                return string.Empty;

            if (obj.GetType() == typeof(string))
                return "\"" + obj.ToString() + "\"";

            return obj.ToString();
        }

        public static string ReEscape(this string s)
        {
            if (s.IndexOfAny(new char[] { '\0', '\n', '\r', '\b', '\a', '\f', '\t', '\v' }) >= 0)
            {
                s = s.Replace("\0", "\\0");
                s = s.Replace("\n", "\\n");
                s = s.Replace("\r", "\\r");
                s = s.Replace("\b", "\\b");
                s = s.Replace("\a", "\\a");
                s = s.Replace("\f", "\\f");
                s = s.Replace("\t", "\\t");
                s = s.Replace("\v", "\\v");
            }
            return "\"" + s + "\"";
        }

        public static string[] GetNames(this FuncDesc fd, ITypeInfo ti)
        {
            var names = new string[fd.elemdescParams.Length + 1];
            ti.GetNames(fd.memid, names, fd.elemdescParams.Length + 1, out int pcNames);
            return names;
        }
    }
}

