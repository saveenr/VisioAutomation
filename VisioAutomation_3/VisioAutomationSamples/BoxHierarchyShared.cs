using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomationSamples
{
    public class BoxHierarchyShared
    {
        public static void DrawBoxHierarchyDrawNode(
            VisioAutomation.Layout.BoxHierarchy.Node<IVisio.Shape> node,
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

        public static VisioAutomation.Layout.BoxHierarchy.BoxHierarchyLayout<IVisio.Shape>
            CreateSampleBoxHierarchyLayout()
        {
            // Create a new layout
            var layout =
                new VisioAutomation.Layout.BoxHierarchy.BoxHierarchyLayout<IVisio.Shape>(
                    VisioAutomation.Layout.BoxHierarchy.LayoutDirection.Vertical);

            // Add the nodes and specify their sizes and in what direction to draw them
            var g0 = layout.Root;
            g0.AlignmentHorizontal = VA.Drawing.AlignmentHorizontal.Right;
            g0.Padding = 0.5;

            var g1 = g0.AddNode(VisioAutomation.Layout.BoxHierarchy.LayoutDirection.Vertical);
            g1.AlignmentHorizontal = VA.Drawing.AlignmentHorizontal.Center;
            g1.Padding = 0.25;
            g1.ChildSeparation = 0.25;
            g1.AddNode(1, 0.25);
            g1.AddNode(1.25, 0.25);
            g1.AddNode(1.50, 0.25);
            g1.AddNode(1.75, 0.25);
            g1.AddNode(2, 0.25);

            var g2 = g0.AddNode(VisioAutomation.Layout.BoxHierarchy.LayoutDirection.Horizontal);
            g2.AlignmentVertical = VA.Drawing.AlignmentVertical.Center;
            g2.Padding = 0.10;
            g2.ChildSeparation = 0.05;
            g2.AddNode(0.25, 0.26, VA.Drawing.AlignmentVertical.Top);
            g2.AddNode(3.5, 0.5, VA.Drawing.AlignmentVertical.Center);
            g2.AddNode(0.5, 0.5);
            g2.AddNode(0.5, 0.6);
            g2.AddNode(0.5, 0.7);
            g2.AddNode(0.5, 0.8);

            var g3 = g2.AddNode(VisioAutomation.Layout.BoxHierarchy.LayoutDirection.Vertical);
            g3.Padding = 0.25;
            g3.ChildSeparation = 0.20;
            g3.AddNode(0.30, 0.25, VA.Drawing.AlignmentHorizontal.Right);
            g3.AddNode(0.25, 0.25, VA.Drawing.AlignmentHorizontal.Center);
            g3.AddNode(0.20, 0.25);
            g3.AddNode(0.15, 0.25);

            return layout;
        }
    }
}