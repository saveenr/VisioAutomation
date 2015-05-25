using System.Windows.Forms;

namespace VisioAutomation.UI
{
    /// http://support.microsoft.com/kb/318804
    /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
    /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookstructures/cwpstruct.asp
    /// http://www.codeproject.com/KB/cs/globalhook.aspx
    
    public class Hook
    {
        public delegate void KeyboardDelegate(KeyEventArgs e);

        public KeyboardDelegate OnKeyDown;
        private int m_hHook = 0;
        private NativeMethods.HookProc m_HookCallback;

        public void SetHook(bool enable)
        {
            if (enable && this.m_hHook == 0)
            {
                this.m_HookCallback = this.HookCallbackProc;
                var module = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0];
                this.m_hHook = NativeMethods.SetWindowsHookEx(WinUtil.WH_KEYBOARD_LL, this.m_HookCallback,
                                                                     System.Runtime.InteropServices.Marshal.GetHINSTANCE
                                                                         (module),
                                                                     0);
                if (this.m_hHook == 0)
                {
                    MessageBox.Show(
                        "SetHook Failed. Please make sure the 'Visual Studio Host Process' on the debug setting page is disabled");
                    return;
                }
                return;
            }

            if (enable == false && this.m_hHook != 0)
            {
                NativeMethods.UnhookWindowsHookEx(this.m_hHook);
                this.m_hHook = 0;
            }
        }

        private int HookCallbackProc(int nCode, System.IntPtr wParam, System.IntPtr lParam)
        {
            if (nCode < 0)
            {
                return NativeMethods.CallNextHookEx(this.m_hHook, nCode, wParam, lParam);
            }
            else
            {
                //Marshall the data from the callback.
                var hookstruct = (WinUtil.KeyboardHookStruct) System.Runtime.InteropServices.Marshal.PtrToStructure(lParam, typeof(WinUtil.KeyboardHookStruct));

                if (this.OnKeyDown != null && wParam.ToInt32() == WinUtil.WM_KEYDOWN)
                {
                    var key = (Keys) hookstruct.vkCode;
                    const Keys shift = Keys.Shift;
                    const Keys control = Keys.Control;
                    Keys modkeys = Control.ModifierKeys;

                    if ((modkeys & shift) == shift)
                    {
                        key |= shift;
                    }

                    if ((modkeys & control) == control)
                    {
                        key |= control;
                    }

                    var e = new KeyEventArgs(key);
                    e.Handled = false;
                    this.OnKeyDown(e);

                    if (e.Handled)
                    {
                        return 1;
                    }
                }

                int result = 0;
                if (this.m_hHook != 0)
                {
                    result = NativeMethods.CallNextHookEx(this.m_hHook, nCode, wParam, lParam);
                }

                return result;
            }
        }
    }
}