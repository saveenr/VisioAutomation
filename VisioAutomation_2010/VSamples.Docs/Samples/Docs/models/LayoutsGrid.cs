using GRID = VisioAutomation.Models.Layouts.Grid;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/layouts-grid.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsLayoutsGridSamples
    {
        public static void Sample_01(IVisio.Page visioPage, IVisio.Master rectMaster)
        {
            // <sample id="models/layouts-grid#01" doc="models/layouts-grid.md" line="17">
            int cols = 3;
            int rows = 6;
            var cellsize = new VA.Core.Size(0.5, 0.25);

            var grid = new GRID.GridLayout(cols, rows, cellsize, rectMaster);
            grid.Origin = new VA.Core.Point(0, 4);
            grid.PerformLayout();
            grid.Render(visioPage);
            // </sample>
        }
    }
}
