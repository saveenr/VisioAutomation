using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using BoxL = VisioAutomation.Layout.BoxLayout;

namespace VisioAutomationSamples
{
    public class BoxLayoutShared
    {
        public static void DrawNode(
            BoxL.Node node,
            VA.Drawing.Rectangle rect, IVisio.Page page)
        {
            var src_fillfg = VA.ShapeSheet.SRCConstants.FillForegnd;
            var src_linecolor = VA.ShapeSheet.SRCConstants.LineColor;
            
            var shape = page.DrawRectangle(rect);
            node.Data = shape;
            if (node.ChildCount > 0)
            {
                var cell_linecolor = shape.GetCell(src_linecolor);
                cell_linecolor.FormulaU = "rgb(255,0,0)";
            }
            else
            {
                var cell_fillfg = shape.GetCell(src_fillfg);
                cell_fillfg.FormulaU = "rgb(240,240,240)";
            }
        }

        public static BoxL.BoxLayout
            CreateSampleLayout()
        {
            // Create a new layout
            var layout =
                new BoxL.BoxLayout(BoxL.LayoutDirection.Vertical);

            // Add the nodes and specify their sizes and in what direction to draw them
            var g0 = layout.Root;
            g0.AlignmentHorizontal = VA.Drawing.AlignmentHorizontal.Right;
            g0.Padding = 0.5;

            var g1 = g0.AddColumn();
            g1.AlignmentHorizontal = VA.Drawing.AlignmentHorizontal.Center;
            g1.Padding = 0.25;
            g1.ChildSeparation = 0.25;
            g1.AddBox(1, 0.25);
            g1.AddBox(1.25, 0.25);
            g1.AddBox(1.50, 0.25);
            g1.AddBox(1.75, 0.25);
            g1.AddBox(2, 0.25);

            var g2 = g0.AddRow();
            g2.AlignmentVertical = VA.Drawing.AlignmentVertical.Center;
            g2.Padding = 0.10;
            g2.ChildSeparation = 0.05;
            g2.AddRow(0.25, 0.26, VA.Drawing.AlignmentVertical.Top);
            g2.AddRow(3.5, 0.5, VA.Drawing.AlignmentVertical.Center);
            g2.AddBox(0.5, 0.5);
            g2.AddBox(0.5, 0.6);
            g2.AddBox(0.5, 0.7);
            g2.AddBox(0.5, 0.8);

            var g3 = g2.AddColumn();
            g3.Padding = 0.25;
            g3.ChildSeparation = 0.20;
            g3.AddColumn(0.30, 0.25, VA.Drawing.AlignmentHorizontal.Right);
            g3.AddColumn(0.25, 0.25, VA.Drawing.AlignmentHorizontal.Center);
            g3.AddBox(0.20, 0.25);
            g3.AddBox(0.15, 0.25);

            return layout;
        }
    }
}