/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;

namespace olewoo
{
    class IDLHelperTab : IDisposable
    {
        IDLFormatter _ih;

        public IDLHelperTab(IDLFormatter ih)
        {
            ih.Indent();
            _ih = ih;
        }

        public void Dispose() => _ih.Dedent();
    }
}
