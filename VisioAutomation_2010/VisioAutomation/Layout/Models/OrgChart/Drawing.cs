using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;
using VAL = VisioAutomation.Layout;

namespace VisioAutomation.Layout.Models.OrgChart
{
    public class Drawing
    {
        public Node Root { get; set; }

        public void Render(IVisio.Application app)
        {
            var layout = new OrgChartLayout();
            layout.RenderToVisio(this, app);
        }
    }
}