using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace VSamples.Samples.Misc
{
    public class GridOfMasters : SampleMethodBase
    {
        public override void Run()
        {
            // http://blogs.msdn.com/saveenr/archive/2008/08/06/visioautoext-simplifying-dropmany-to-quickly-draw-a-grid.aspx

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Resize the page to a sqaure
            var page_size = new VA.Core.Size(4, 4);
            SampleEnvironment.SetPageSize(page, page_size);

            // Load the Stencil
            var application = page.Application;
            var documents = application.Documents;
            var stencil = documents.OpenStencil("basic_u.vss");
            var stencil_masters = stencil.Masters;
            var master = stencil_masters["Rectangle"];

            // Calculate where to drop the masters
            int num_cols = 10;
            int num_rows = 10;

            var centerpoints = new List<VA.Core.Point>(num_rows * num_cols);
            foreach (var row in Enumerable.Range(0, num_rows))
            {
                foreach (var col in Enumerable.Range(0, num_cols))
                {
                    var p = new VA.Core.Point(row * 1.0, col * 1.0);
                    centerpoints.Add(p);
                }
            }

            var masters = new[] {master};

            // Draw the masters
            var shapeids = page.DropManyU(masters, centerpoints);

            var bordersize = new VA.Core.Size(1, 1);
            page.ResizeToFitContents(bordersize);
        }
    }
}