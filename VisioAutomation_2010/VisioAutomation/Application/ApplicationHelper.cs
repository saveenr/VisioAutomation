using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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

        public static string GetContentLocation(Microsoft.Office.Interop.Visio.Application app)
        {
            var ver = GetVersion(app);

            if (ver.Major == 14)
            {
                string path = System.IO.Path.Combine(app.Path, "Visio Content");
                path = System.IO.Path.Combine(path, app.Language.ToString(System.Globalization.CultureInfo.InvariantCulture));
                return path;
            }

            if (ver.Major >= 15)
            {
                string path = System.IO.Path.Combine(app.Path, "Visio Content");
                path = System.IO.Path.Combine(path, app.Language.ToString(System.Globalization.CultureInfo.InvariantCulture));
                return path;

            }

            throw new System.ArgumentException("This version of visio not supported");
        }

        public static string GetXMLErrorLogFilename(Microsoft.Office.Interop.Visio.Application app)
        {
            // the location of the xml error log file is specific to the user
            // we need to retrieve it from the registry
            var hkcu = Microsoft.Win32.Registry.CurrentUser;

            // The reg path is specific to the version of visio being used

            string ver = app.Version;
            string ver_normalized = ver.Replace(",", ".");

            string path = string.Format(@"Software\Microsoft\Office\{0}\Visio\Application", ver_normalized);

            string logfilename = null;
            using (var key_visio_application = hkcu.OpenSubKey(path))
            {
                if (key_visio_application == null)
                {
                    // key doesn't exist - can't continue
                    throw new AutomationException("Could not find the key visio application key in hkcu");
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
    }
}