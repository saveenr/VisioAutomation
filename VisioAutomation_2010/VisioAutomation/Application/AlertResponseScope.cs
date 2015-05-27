using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Application
{
    /// <summary>
    /// Using an IDisposable pattern, this class allows the Visio Application's AlertResponse 
    /// property to be set with a guarantee that the AlertResponse will be set back to its
    /// previous value. 
    /// </summary>
    public class AlertResponseScope : System.IDisposable
    {
        private readonly AlertResponseCode old_alertresponse;
        private readonly IVisio.Application app;

        public AlertResponseScope(IVisio.Application app, AlertResponseCode value)
        {
            this.app = app;
            this.old_alertresponse = (AlertResponseCode)this.app.AlertResponse;
            this.app.AlertResponse = (short)value;
        }

        public void Dispose()
        {
            this.app.AlertResponse = (short)this.old_alertresponse;
        }
    }
}