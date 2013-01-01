using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;
using VAL = VisioAutomation.Layout;

namespace VisioAutomation.Layout.Models.OrgChart
{
    public class Drawing
    {
        public List<Node> Roots { get; set; }

        public Drawing()
        {
            this.Roots = new List<Node>();
        }

        public void Render(IVisio.Application app)
        {
            var layout = new OrgChartLayout();
            layout.RenderToVisio(this, app);
        }
    }
}