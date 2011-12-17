namespace VisioAutomation.UI
{
    public static class ColorPicker_NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern System.IntPtr GetWindowLong32(System.IntPtr hWnd, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern System.IntPtr SetWindowLongPtr32(System.IntPtr hWnd, int nIndex, int dwNewLong);

        public delegate int HookProc(int nCode, System.IntPtr wParam, System.IntPtr lParam);

        /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, System.IntPtr hInstance, int threadId);

        /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, System.IntPtr wParam, System.IntPtr lParam);
    }
}