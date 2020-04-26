/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using System.Text;
using olewoo.interop;

namespace olewoo
{
    class IDLGrabber : interop.IDLFormatter
    {
        StringBuilder _s = new StringBuilder();
        public override void AddLink(string s, string s2) => _s.Append(s);
        public override void AddString(string s) => _s.Append(s);
        public string Value => _s.ToString();

        public override string ToString() => Value;
    }
}

