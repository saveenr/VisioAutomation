namespace VisioAutomation.Application
{
    /// <summary>
    /// Using an IDisposable pattern, this class allows the Visio Application's AlertResponse 
    /// property to be set with a guarantee that the AlertResponse will be set back to its
    /// previous value. 
    /// </summary>
    public class AlertResponseScope : System.IDisposable
    {
        private readonly AlertResponseCode _old_alertresponse;
        private readonly IVisio.Application _app;

        public AlertResponseScope(IVisio.Application app, AlertResponseCode value)
        {
            this._app = app;
            this._old_alertresponse = (AlertResponseCode)this._app.AlertResponse;
            this._app.AlertResponse = (short)value;
        }

        public void Dispose()
        {
            this._app.AlertResponse = (short)this._old_alertresponse;
        }
    }
}