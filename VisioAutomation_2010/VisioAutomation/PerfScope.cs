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
        private bool configured;
        private IVisio.Application app;
        private PerfSettings old_settings;

        public PerfScope(IVisio.Application vis)
        {
            configured = false;
            this.app = vis;
            this.begin_config();
        }

        private void begin_config()
        {
            save_old_config();
            configured = true;
            set_config();
        }

        private void end_config()
        {
            if (configured)
            {
                restore_config();
                configured = false;
            }
        }

        private void set_config()
        {
            var fast = new PerfSettings();
            fast.DeferRecalc = 1;
            fast.ScreenUpdating = 0;
            fast.EnableAutoConnect = false;
            fast.LiveDynamics = false;

            fast.Apply(app);
        }

        private void save_old_config()
        {
            this.old_settings = new PerfSettings();
            this.old_settings.Load(this.app);
        }

        private void restore_config()
        {
            this.old_settings.Apply(this.app);
        }

        public void Dispose()
        {
            this.end_config();
        }
    }
}