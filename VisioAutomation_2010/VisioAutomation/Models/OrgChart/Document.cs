using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.OrgChart
{
    public class Document
    {
        public List<Node> OrgCharts { get; private set; }

        public Document()
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