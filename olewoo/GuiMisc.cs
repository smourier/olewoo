using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace olewoo
{
    delegate void EndUpdateDelg();

    class UpdateSuspender : IDisposable
    {
        readonly EndUpdateDelg _eud;
        public UpdateSuspender(TreeView t)
        {
            t.BeginUpdate();
            _eud = t.EndUpdate;
        }

        public void Dispose() => _eud();
    }

    class TBUpdateSuspender : IDisposable
    {
        [DllImport("user32")]
        private extern static IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp); 
        
        TextBoxBase _tbb;
        public TBUpdateSuspender(TextBoxBase tbb)
        {
            _tbb = tbb;
            SendMessage(_tbb.Handle, 0xb, (IntPtr)0, IntPtr.Zero); 
        }

        public void Dispose()
        {
            SendMessage(_tbb.Handle, 0xb, (IntPtr)1, IntPtr.Zero);
            _tbb.Invalidate();
        }
    }
}
