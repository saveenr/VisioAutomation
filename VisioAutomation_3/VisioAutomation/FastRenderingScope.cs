using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public class FastRenderingScope : System.IDisposable
    {
        private bool old_autoconnect;
        private bool old_livedynamics;
        private short old_screenupdating;
        private short old_deferrecalc;
        private bool configured;
        private IVisio.Application app;

        const short screen_updating_fast = 0; // disable screen updating
        const short defer_recalc_fast = 1; // defer recalc
        const bool enable_autoconnect_fast = false; // diable autoconnect
        const bool livedynamics_fast = false; // diable live dynamics

        public FastRenderingScope(IVisio.Application vis)
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
            var app_settings = app.Settings;
            app.ScreenUpdating = screen_updating_fast;
            app.DeferRecalc = defer_recalc_fast;
            app_settings.EnableAutoConnect = enable_autoconnect_fast;
            app.LiveDynamics = livedynamics_fast;
        }

        private void save_old_config()
        {
            var app_settings = app.Settings;

            this.old_livedynamics = app.LiveDynamics;
            this.old_autoconnect = app_settings.EnableAutoConnect;
            this.old_deferrecalc = app.DeferRecalc;
            this.old_screenupdating = app.ScreenUpdating;
        }

        private void restore_config()
        {
            app.ScreenUpdating = this.old_screenupdating;
            app.DeferRecalc = this.old_deferrecalc;
            app.Settings.EnableAutoConnect = this.old_autoconnect;
            app.LiveDynamics = this.old_livedynamics;
        }

        public void Dispose()
        {
            this.end_config();
        }
    }
}