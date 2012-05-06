using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    internal class PerfSettings
    {
        public bool? EnableAutoConnect;
        public bool? LiveDynamics;
        public short? ScreenUpdating;
        public short? DeferRecalc;        

        public PerfSettings()
        {
            
        }

        public void Load(IVisio.Application app)
        {
            var app_settings = app.Settings;
            this.LiveDynamics = app.LiveDynamics;
            this.EnableAutoConnect = app_settings.EnableAutoConnect;
            this.DeferRecalc = app.DeferRecalc;
            this.ScreenUpdating = app.ScreenUpdating;
        }

        public void Apply(IVisio.Application app)
        {
            if (this.ScreenUpdating.HasValue) {app.ScreenUpdating = this.ScreenUpdating.Value;}
            if (this.DeferRecalc.HasValue) {app.DeferRecalc = this.DeferRecalc.Value;}
            if (this.EnableAutoConnect.HasValue) {app.Settings.EnableAutoConnect = this.EnableAutoConnect.Value;}
            if (this.LiveDynamics.HasValue) {app.LiveDynamics = this.LiveDynamics.Value;}
        }

    }

    public class PerfScope : System.IDisposable
    {
        private IVisio.Application app;
        private PerfSettings old_settings;

        public PerfScope(IVisio.Application vis)
        {
            this.app = vis;

            // save the old settings
            this.old_settings = new PerfSettings();
            this.old_settings.Load(this.app);

            // Set the new settings
            var fast = new PerfSettings();
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