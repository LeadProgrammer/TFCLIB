using System;
using System.Collections.Generic;
using System.Text;
//using System.Windows.Forms;

namespace LIBDI
{
    public class MicroSoftWindows
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        //Win32 API calls necesary to raise an unowned processs main window
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        public MicroSoftWindows()
        {
            //
            // TODO: Add constructor logic here
            //
        }


        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            public IntPtr _hwnd;

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
            public WindowWrapper()
            {
                IntPtr handle = GetForegroundWindow();
                _hwnd = handle;
            }
        }
    }
}
