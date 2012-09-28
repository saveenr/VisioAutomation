using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using BoxL = VisioAutomation.Layout.Models.BoxLayout;

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
            /*
            if (node.Count > 0)
            {
                var cell_linecolor = shape.GetCell(src_linecolor);
                cell_linecolor.FormulaU = "rgb(255,0,0)";
            }
            else
            {
                var cell_fillfg = shape.GetCell(src_fillfg);
                cell_fillfg.FormulaU = "rgb(240,240,240)";
            }*/
        }
    }
}