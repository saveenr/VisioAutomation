using System;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Application
{
    public static class ApplicationHelper
    {
        /// <summary>
        /// Finds running instances of Visio
        /// </summary>
        /// <remarks>
        /// On occasion, despite an instance of visio running, this method will still return null.</remarks>
        /// <returns>null if an instance cannot be found, otherwise returns the instance</returns>
        public static IVisio.Application FindRunningApplication()
        {
            const string progid = VA.Internal.Constants.VisioApplication_ProgID;
            object o = null;

            try
            {
                o = System.Runtime.InteropServices.Marshal.GetActiveObject(progid);

            }
            catch (System.Runtime.InteropServices.COMException exc)
            {
                // if you are wondering why the conversion to uint is needed below
                // http://stackoverflow.com/questions/1426147/catching-comexception-specific-error-code

                const uint MK_E_UNAVAILABLE = 0x800401E3;
                if (((uint)exc.ErrorCode) == MK_E_UNAVAILABLE) // MK_E_UNAVAILABLE
                {
                    return null;
                }
            }

            var app = (IVisio.Application) o;
            return app;
        }

        public static void Quit(IVisio.Application app, bool force_close)
        {
            short old = app.AlertResponse;
            if (force_close)
            {
                const short new_alert_response = 7;
                app.AlertResponse = new_alert_response;
            }

            app.Quit();
        }       
        
        public static void BringWindowToTop(IVisio.Application app)
        {
            var visio_window_handle = new System.IntPtr(app.WindowHandle32);
            VA.Internal.Interop.NativeMethods.BringWindowToTop(visio_window_handle);
        }

        private static ApplicationInformation _app_info;

        public static ApplicationInformation GetInformation(IVisio.Application app)
        {
            _app_info = _app_info ?? new ApplicationInformation(app);
            return _app_info;
        }
    }
}