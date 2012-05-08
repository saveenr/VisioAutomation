using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation
{
    public class PerfScope : System.IDisposable
    {
        private IVisio.Application app;
        private VA.Internal.PerfSettings old_settings;

        public PerfScope(IVisio.Application vis)
        {
            this.app = vis;

            // save the old settings
            this.old_settings = new VA.Internal.PerfSettings();
            this.old_settings.Load(this.app);

            // Set the new settings
            var fast = new VA.Internal.PerfSettings();
            fast.DeferRecalc = 1;
            fast.ScreenUpdating = 0;
            fast.EnableAutoConnect = false;
            fast.LiveDynamics = false;
            fast.Apply(app);
        }

        public void Dispose()
        {
            this.old_settings.Apply(this.app);
        }
    }
}