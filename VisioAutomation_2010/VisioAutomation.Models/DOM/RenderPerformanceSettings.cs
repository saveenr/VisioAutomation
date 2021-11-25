

namespace VisioAutomation.Models.Dom
{
    public class RenderPerformanceSettings
    {
        public bool? EnableAutoConnect;
        public bool? LiveDynamics;
        public short? ScreenUpdating;
        public short? DeferRecalc;

        public RenderPerformanceSettings()
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
            var app_settings = app.Settings;
            if (this.ScreenUpdating.HasValue) { app.ScreenUpdating = this.ScreenUpdating.Value; }
            if (this.DeferRecalc.HasValue) { app.DeferRecalc = this.DeferRecalc.Value; }
            if (this.EnableAutoConnect.HasValue) { app_settings.EnableAutoConnect = this.EnableAutoConnect.Value; }
            if (this.LiveDynamics.HasValue) { app.LiveDynamics = this.LiveDynamics.Value; }
        }

    }
}