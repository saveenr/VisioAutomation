using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;
using VAL = VisioAutomation.Layout;

namespace VisioAutomation.Layout.OrgChart
{
    public class Drawing
    {
        public Node Root { get; set; }

        public void Render(IVisio.Application app)
        {
            var renderer = new OrgChartLayout();
            renderer.RenderToVisio(this, app);
        }
    }
}