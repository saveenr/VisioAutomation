using System.Linq;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

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

        public static string GetContentLocation(IVisio.Application app)
        {
            var ver = ApplicationHelper.GetVersion(app);
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            string app_Lang = app.Language.ToString(invariant_culture);
            var str_visio_content = "Visio Content";

            if (ver.Major == 14)
            {
                string path = System.IO.Path.Combine(app.Path, str_visio_content);
                path = System.IO.Path.Combine(path, app_Lang);
                return path;
            }

            if (ver.Major >= 15)
            {
                string path = System.IO.Path.Combine(app.Path, str_visio_content);
                path = System.IO.Path.Combine(path, app_Lang);
                return path;
            }

            string msg = string.Format(invariant_culture,"VisioAutomation does not support Visio version {0}", ver.Major);
            throw new System.ArgumentException(msg);
        }

        public static string GetXmlErrorLogFilename(IVisio.Application app)
        {
            // the location of the xml error log file is specific to the user
            // we need to retrieve it from the registry
            var hkcu = Microsoft.Win32.Registry.CurrentUser;

            // The reg path is specific to the version of visio being used

            string ver = app.Version;
            string ver_normalized = ver.Replace(",", ".");

            string path = $@"Software\Microsoft\Office\{ver_normalized}\Visio\Application";

            string logfilename = null;
            using (var key_visio_application = hkcu.OpenSubKey(path))
            {
                if (key_visio_application == null)
                {
                    // key doesn't exist - can't continue
                    throw new InternalAssertionException("Could not find the key visio application key in hkcu");
                }

                var subkeynames = key_visio_application.GetValueNames();
                if (!subkeynames.Contains("XMLErrorLogName"))
                {
                    return null;
                }

                logfilename = (string)key_visio_application.GetValue("XMLErrorLogName");
            }

            // the folder that contains the file is located in the users internet cache
            // C:\Users\<your alias>\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.MSO\VisioLogFiles
            string internetcache = System.Environment.GetFolderPath(System.Environment.SpecialFolder.InternetCache);
            string folder = System.IO.Path.Combine(internetcache, @"Content.MSO\VisioLogFiles");

            var s = System.IO.Path.Combine(folder, logfilename);
            System.Diagnostics.Debug.WriteLine("XmlErrorLogFilename: " + s);

            return s;
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
        
        public static void BringWindowToTop(IVisio.Application app)
        {
            var visio_window_handle = new System.IntPtr(app.WindowHandle32);
            Internal.Interop.NativeMethods.BringWindowToTop(visio_window_handle);
        }
    }
}