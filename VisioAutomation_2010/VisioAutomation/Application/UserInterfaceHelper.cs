namespace VisioAutomation.Application
{
    public static class UserInterfaceHelper
    {
        /// <summary>
        /// Allows a windows form to be used as the UI for an anchor window
        /// </summary>
        public static void AttachWindowsForm(
            IVisio.Window anchor_window,
            System.Windows.Forms.Form the_form)
        {
            if (anchor_window == null)
            {
                throw new System.ArgumentNullException(nameof(anchor_window));
            }

            if (the_form == null)
            {
                throw new System.ArgumentNullException(nameof(the_form));
            }

            // Show the form as a modeless dialog.
            the_form.Show();

            // Get the window handle of the form.
            int hwnd = the_form.Handle.ToInt32();
            var hwnd_as_intptr = new System.IntPtr(hwnd);

            // Set the window properties to make it a visible child window.
            const int window_prop_index = VisioAutomation.Internal.NativeMethods.GWL_STYLE;
            const int window_prop_value = VisioAutomation.Internal.NativeMethods.WS_CHILD | VisioAutomation.Internal.NativeMethods.WS_VISIBLE;
            VisioAutomation.Internal.NativeMethods.SetWindowLong(hwnd_as_intptr, window_prop_index, window_prop_value);

            // Set the anchor bar window as the parent of the form.
            VisioAutomation.Internal.NativeMethods.SetParent(hwnd, anchor_window.WindowHandle32);

            // Force a resize of the anchor bar so it will refresh.
            int left, top, width, height;
            anchor_window.GetWindowRect(out left, out top, out width, out height);
            anchor_window.SetWindowRect(left, top, width - 1, height - 1);
            anchor_window.SetWindowRect(left, top, width, height);

            // Set the dock property of the form to fill, so that the form
            // automatically resizes to the size of the anchor bar.
            the_form.Dock = System.Windows.Forms.DockStyle.Fill;

            // had to set to false to prevent a resizing problem (it was originally set to true)
            the_form.AutoSize = true;
        }

        /// <summary>
        /// Creates a new anchor window
        /// </summary>
        public static IVisio.Window AddAnchorWindow(IVisio.Window parent_window,
                                                    string caption,
                                                    object window_states,
                                                    object window_types,
                                                    System.Drawing.Rectangle rect)
        {
            if (parent_window == null)
            {
                throw new System.ArgumentNullException(nameof(parent_window));
            }

            var parents_windows = parent_window.Windows;
            var anchor_window = parents_windows.Add(
                caption,
                window_states,
                window_types,
                rect.Left,
                rect.Top,
                rect.Width,
                rect.Height,
                0, 0, 0);
            return anchor_window;
        }
    }
}