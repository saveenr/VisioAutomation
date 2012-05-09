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
        private AlertResponseCode old_alertresponse;
        private bool configured = false;
        private readonly IVisio.Application app;
        private readonly AlertResponseCode conf_val;

        internal AlertResponseScope(IVisio.Application app, AlertResponseCode value)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException("app");
            }

            this.configured = false;
            this.app = app;
            this.conf_val = value;
            this.begin_config();
        }

        private void begin_config()
        {
            this.old_alertresponse = (AlertResponseCode) this.app.AlertResponse;
            this.app.AlertResponse = (short) this.conf_val;
            this.configured = true;
        }

        private void end_config()
        {
            if (this.configured)
            {
                this.app.AlertResponse = (short) this.old_alertresponse;
                this.configured = false;
            }
        }

        public void Dispose()
        {
            this.end_config();
        }
    }
}