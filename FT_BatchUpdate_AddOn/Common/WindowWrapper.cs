using System;
using System.Collections.Generic;

namespace FTS.Common
{
    class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;

        // Property
        public virtual IntPtr Handle
        {
            get { return _hwnd; }
        }

        // Constructor
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
    }
}
