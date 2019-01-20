/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using olewoo.interop;

namespace olewoo
{
    public abstract class IDLFormatter : IDLFormatter_iop
    {
        protected int _tabdepth;

        public void Indent() => _tabdepth++;
        public void Dedent() => _tabdepth--;
        public abstract void NewLine();

        public void AppendLine(string s)
        {
            AddString(s);
            NewLine();
        }
    }
}
