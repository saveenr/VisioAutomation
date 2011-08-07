using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using IG = VisioAutomation.Infographics;

namespace VisioAutomationSamples
{
    public static class InfographcisSamples2
    {
        public static void PieChartGrid()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var data= new[]
                                 {
                                     new {value = 12.0,cat="Prototype"},
                                     new {value = 11.0,cat="YUI"},
                                     new {value = 16.0,cat="JQuery UI"},
                                     new {value = 38.0,cat="JQuery"},
                                     new {value = 13.0,cat="MooTools"},
                                     new {value = 9.0,cat="Other"}
                                 };

            var datapoints =
                data.Select(d => new IG.DataPoint(d.value, string.Format("{0} ({1})", d.cat, d.value))).ToList();

            var g = new VA.Infographics.SingleValuePieChartGrid();
            g.FontName = "Segoe UI";
            g.DataPoints = datapoints;

            g.Draw(page);

            page.ResizeToFitContents(1,1);
        }
    }
}