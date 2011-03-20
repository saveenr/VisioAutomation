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
                    app = new IVisio.ApplicationClass();
                    var documents = app.Documents;
                    documents.Add("");
                }
                return app;
            }
        }
    }
}