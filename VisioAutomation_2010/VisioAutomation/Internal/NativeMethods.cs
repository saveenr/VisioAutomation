using System;
using System.Runtime.InteropServices;

namespace VisioAutomation.Internal
{
    internal static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        public static string GetWindowText(IntPtr hWnd)
        {
            // Allocate correct string length first
            int length = GetWindowTextLength(hWnd);
            var sb = new System.Text.StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }
        
        [DllImport("user32.dll")]
        public static extern uint RealGetWindowClass(IntPtr hwnd, [Out] System.Text.StringBuilder pszType, uint cchType);

        public const short GWL_STYLE = -16;
        public const int WS_CHILD = 0x40000000;
        public const int WS_VISIBLE = 0x10000000;

        [DllImport("User32.dll")]
        public static extern int SetParent(int hWndChild, int hWndNewParent);

        [DllImport("User32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(IntPtr hWnd);
        
        [DllImport("Ole32.Dll")]
        public static extern int CreateBindCtx(int reserved, out System.Runtime.InteropServices.ComTypes.IBindCtx bind_ctx);
    }
}