

using System;
using System.Linq;


namespace PMExportToPS
{
    class WinFormWrapper : System.Windows.Forms.IWin32Window
    {

        IntPtr _hwnd;
        int _inputedInt;

		public int InputedInt {
			get {
				return _inputedInt;
			}
		}

        public WinFormWrapper(int hwnd)
        {
            _hwnd = new IntPtr(hwnd);
            _inputedInt = hwnd;
        }


        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }
}
