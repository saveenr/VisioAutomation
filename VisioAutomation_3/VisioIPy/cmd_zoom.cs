using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void SetZoom(VA.Scripting.Zoom zoom)
        {
            this.ScriptingSession.View.Zoom( zoom);
        }

        public void SetZoomPercentage(double value)
        {
            this.ScriptingSession.View.ZoomToPercentage(value); 
        }

        public double GetZoom()
        {
            return this.ScriptingSession.View.GetActiveZoom();
        }
    }
}