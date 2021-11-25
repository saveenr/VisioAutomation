
namespace VisioAutomation.Models.Dom;

class RenderPerformanceScope : System.IDisposable
{
    private readonly IVisio.Application _app;
    private readonly RenderPerformanceSettings _old_settings;

    public RenderPerformanceScope(IVisio.Application vis, RenderPerformanceSettings new_settings)
    {
        this._app = vis;

        // save the old settings
        this._old_settings = new RenderPerformanceSettings();
        this._old_settings.Load(this._app);

        // Set the new settings
        new_settings.Apply(this._app);
    }

    public void Dispose()
    {
        this._old_settings.Apply(this._app);
    }
}