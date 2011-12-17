using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomationSamples
{
    public class SampleEnvironment
    {
        private static IVisio.Application app;

        public static IVisio.Application Application
        {
            get
            {
                if (app== null)
                {
                    // there is no application object associated with
                    // this session, so create one
                    create_new_app_instance();
                }
                else
                {
                    // there is an application object associated with this session

                    // before we continue we should try to validate that the
                    // application is valid - the user might have closed the application
                    // leaving us with an application object that is invalid

                    try
                    {
                        // try to do something simple, read-only, and fast with the application object
                        var app_version = app.Version;
                    }
                    catch (System.Runtime.InteropServices.COMException ce)
                    {
                        // If a COMException is thrown, this indicates that the
                        // application object is invalid, so create a new one
                        create_new_app_instance();
                    }                   
                }
                return app;
            }
        }

        private static void create_new_app_instance()
        {
            app = new IVisio.Application();
            var documents = app.Documents;
            documents.Add("");
        }
    }
}