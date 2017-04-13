using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    class RenderPerformanceScope : System.IDisposable
    {
        private readonly IVisio.Application _app;
        private readonly RenderPerforfmanceSettings _old_settings;

        public RenderPerformanceScope(IVisio.Application vis, RenderPerforfmanceSettings new_settings)
        {
            this._app = vis;

            // save the old settings
            this._old_settings = new RenderPerforfmanceSettings();
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