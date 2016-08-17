using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Application
{
    public class PerfScope : System.IDisposable
    {
        private readonly IVisio.Application _app;
        private readonly PerfSettings _old_settings;

        public PerfScope(IVisio.Application vis, PerfSettings new_settings)
        {
            this._app = vis;

            // save the old settings
            this._old_settings = new PerfSettings();
            this._old_settings.Load(this._app);

            // Set the new settings
            new_settings.Apply(this._app);
        }

        public void Dispose()
        {
            this._old_settings.Apply(this._app);
        }
    }
}