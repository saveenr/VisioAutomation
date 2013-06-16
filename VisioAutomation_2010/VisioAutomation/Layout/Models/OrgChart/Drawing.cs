using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;
using VAL = VisioAutomation.Layout;

namespace VisioAutomation.Layout.Models.OrgChart
{
    public class Drawing
    {
        public List<Node> OrgCharts { get; private set; }

        public Drawing()
        {
            this.OrgCharts = new List<Node>();
        }

        public void Render(IVisio.Application app)
        {
            var layout = new OrgChartRenderer();
            layout.RenderToVisio(this, app);
        }
    }
}