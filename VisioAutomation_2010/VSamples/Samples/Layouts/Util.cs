using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.Models.Layouts.Box;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomation.Models;

namespace VSamples.Samples.Layouts
{
    public static class Util
    {
        public static VAM.Layouts.Box.Box AddNodeEx(this VAM.Layouts.Box.Container p, double w, double h, string s)
        {
            var box = p.AddBox(w, h);
            var node_data = new CompareFonts.NodeData();
            node_data.Text = s;
            box.Data = node_data;
            return box;
        }

        public class BoxTwoLevelInfo
        {
            public string Text;
            public bool Render;
            public VisioAutomation.Models.Dom.ShapeCells ShapeCells;
        }

        public static void BoxRender(BoxLayout layout, IVisio.Document doc)
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

            var root_shape = (IVisio.Shape) layout.Root.Data;
            root_shape.CellsU["FillForegnd"].FormulaForceU = "rgb(240,240,240)";
            var margin = new VisioAutomation.Core.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);
        }
    }
}