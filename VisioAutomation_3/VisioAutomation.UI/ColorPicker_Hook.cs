using System.Windows.Forms;

namespace VisioAutomation.UI
{
    /// http://support.microsoft.com/kb/318804
    /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/setwindowshookex.asp
    /// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookstructures/cwpstruct.asp
    /// http://www.codeproject.com/KB/cs/globalhook.aspx
    public class Hook
    {
        public delegate void KeyboardDelegate(System.Windows.Forms.KeyEventArgs e);

        public KeyboardDelegate OnKeyDown;
        private int m_hHook = 0;
        private ColorPicker_NativeMethods.HookProc m_HookCallback;

        public void SetHook(bool enable)
        {
            if (enable && m_hHook == 0)
            {
                m_HookCallback = new ColorPicker_NativeMethods.HookProc(HookCallbackProc);
                System.Reflection.Module module = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0];
                m_hHook = ColorPicker_NativeMethods.SetWindowsHookEx(WinUtil.WH_KEYBOARD_LL, m_HookCallback,
                                                                     System.Runtime.InteropServices.Marshal.GetHINSTANCE
                                                                         (module),
                                                                     0);
                if (m_hHook == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "SetHook Failed. Please make sure the 'Visual Studio Host Process' on the debug setting page is disabled");
                    return;
                }
                return;
            }

            if (enable == false && m_hHook != 0)
            {
                ColorPicker_NativeMethods.UnhookWindowsHookEx(m_hHook);
                m_hHook = 0;
            }
        }

        private int HookCallbackProc(int nCode, System.IntPtr wParam, System.IntPtr lParam)
        {
            if (nCode < 0)
            {
                return ColorPicker_NativeMethods.CallNextHookEx(m_hHook, nCode, wParam, lParam);
            }
            else
            {
                //Marshall the data from the callback.
                WinUtil.KeyboardHookStruct hookstruct =
                    (WinUtil.KeyboardHookStruct)
                    System.Runtime.InteropServices.Marshal.PtrToStructure(lParam, typeof(WinUtil.KeyboardHookStruct));

                if (OnKeyDown != null && wParam.ToInt32() == WinUtil.WM_KEYDOWN)
                {
                    var key = (System.Windows.Forms.Keys) hookstruct.vkCode;
                    const Keys shift = System.Windows.Forms.Keys.Shift;
                    const Keys control = System.Windows.Forms.Keys.Control;
                    Keys modkeys = System.Windows.Forms.Control.ModifierKeys;

                    if ((modkeys & shift) == shift)
                    {
                        key |= shift;
                    }

                    if ((modkeys & control) == control)
                    {
                        key |= control;
                    }

                    var e = new System.Windows.Forms.KeyEventArgs(key);
                    e.Handled = false;
                    OnKeyDown(e);

                    if (e.Handled)
                    {
                        return 1;
                    }
                }

                int result = 0;
                if (m_hHook != 0)
                {
                    result = ColorPicker_NativeMethods.CallNextHookEx(m_hHook, nCode, wParam, lParam);
                }

                return result;
            }
        }
    }
}