using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.Models.Layouts.Box;

namespace VisioAutomationSamples
{
    public static class Util
    {
        public static void Render(BoxLayout layout, Microsoft.Office.Interop.Visio.Document doc)
        {
            layout.PerformLayout();
            var page1 = doc.Pages.Add();
            // and tinker with it
            // render
            var nodes = layout.Nodes.ToList();
            foreach (var node in nodes)
            {
                var shape = page1.DrawRectangle(node.Rectangle);
                node.Data = shape;
            }

            var root_shape = (Microsoft.Office.Interop.Visio.Shape)layout.Root.Data;
            root_shape.CellsU["FillForegnd"].FormulaForceU = "rgb(240,240,240)";
            var margin = new VisioAutomation.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);

        }
    
    }
}