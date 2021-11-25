namespace VisioAutomation.Application
{
    public static class ApplicationHelper
    {
        public static System.Version GetVersion(IVisio.Application app)
        {
            // It's always safer to get the app version via this class because it normalizes the version string
            string verstring = app.Version;
            string verstring_normalized = verstring.Replace(",",".");
            var version = System.Version.Parse(verstring_normalized);
            return version;
        }

        public static void Quit(IVisio.Application app, bool force_close)
        {
            if (force_close)
            {
                const short new_alert_response = 7;
                app.AlertResponse = new_alert_response;
            }

            app.Quit();
        }

        public static string GetContentLocation(IVisio.Application app)
        {
            var ver = ApplicationHelper.GetVersion(app);
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            string app_lang = app.Language.ToString(culture);
            var str_visio_content = "Visio Content";

            if (ver.Major == 14)
            {
                string path = System.IO.Path.Combine(app.Path, str_visio_content);
                path = System.IO.Path.Combine(path, app_lang);
                return path;
            }

            if (ver.Major >= 15)
            {
                string path = System.IO.Path.Combine(app.Path, str_visio_content);
                path = System.IO.Path.Combine(path, app_lang);
                return path;
            }

            string msg = string.Format(culture,"VisioAutomation does not support Visio version {0}", ver.Major);
            throw new System.ArgumentException(msg);
        }
        
        public static void BringWindowToTop(IVisio.Application app)
        {
            var visio_window_handle = new System.IntPtr(app.WindowHandle32);
            VisioAutomation.Internal.NativeMethods.BringWindowToTop(visio_window_handle);
        }
    }
}