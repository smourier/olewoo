/**************************************
 *
 * Part of OLEWOO - http://www.benf.org (2010-2012)
 *
 * CopyLeft, but please credit.
 *
 */
using olewoo.interop;

namespace olewoo
{
    class IDLGrabber : IDLFormatter_iop
    {
        string _s;
        public override void AddLink(string s, string s2) => _s += s;
        public override void AddString(string s) => _s += s;
        public string Value => _s;

        public override string ToString() => _s;
    }
}

