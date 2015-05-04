using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Application
{
    public class PerfScope : System.IDisposable
    {
        private readonly IVisio.Application app;
        private readonly PerfSettings old_settings;

        public PerfScope(IVisio.Application vis, PerfSettings new_settings)
        {
            this.app = vis;

            // save the old settings
            this.old_settings = new PerfSettings();
            this.old_settings.Load(this.app);

            // Set the new settings
            new_settings.Apply(this.app);
        }

        public void Dispose()
        {
            this.old_settings.Apply(this.app);
        }
    }
}