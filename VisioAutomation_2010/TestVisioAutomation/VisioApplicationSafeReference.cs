using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    public class VisioApplicationSafeReference
    {
        // this class ensures that a valid application instance is always available

        private IVisio.Application app;

        public IVisio.Application GetVisioApplication()
        {
            if (this.app == null)
            {
                // obviously if the current app reference is empty
                // then we must create a new once
                this.app = new IVisio.Application();
            }
            else
            {
                // OK, we have an instance, but it may not be valid
                // any longer because someome closed the app
                try
                {
                    // Try doing *something* with the instance
                    string s = this.app.Name;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // This COMException is a hint something
                    // is wrong with the instance. So, create a new
                    // visio application
                    this.app = new IVisio.Application();
                }
            }

            return this.app;
        }
    }
}