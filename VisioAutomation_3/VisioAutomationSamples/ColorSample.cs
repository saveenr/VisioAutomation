using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class ColorSample
    {
        public static void ColorGrid()
        {
            var fill_foregnd = VisioAutomation.ShapeSheet.SRCConstants.FillForegnd;

            int[] vista_desktop_colors = {
                                             0x0A3B76, 0x4395D1, 0x99D9EA, 0x0D686B, 0x00A99D, 0x7ACCC8, 0x82CA9C,
                                             0x74A402,
                                             0xC4DF9B, 0xD9D56F, 0xFFF468, 0xFFF799, 0xFFC20E, 0xEB6119, 0xFBAF5D,
                                             0xE57300, 0xC14000, 0xB82832, 0xD85171, 0xFEDFEC, 0x563F7F, 0xA186BE,
                                             0xD9CFE5
                                         };

            var color_formulas = vista_desktop_colors.Select(x => new VA.Drawing.ColorRGB(x).ToFormula()).ToList();

            const int num_cols = 5;
            const int num_rows = 5;

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            page.SetSize(10, 10);

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];

            var layout = new VA.Layout.Grid.GridLayout(num_cols, num_rows, new VA.Drawing.Size(1, 1), master);
            layout.Origin = new VA.Drawing.Point(0, 0);
            layout.CellSpacing = new VA.Drawing.Size(0, 0);
            layout.RowDirection = VA.Layout.Grid.RowDirection.BottomToTop;

            layout.PerformLayout();
            layout.Render(page);

            int i = 0;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            foreach (var node in layout.Nodes)
            {
                var shapeid = node.ShapeID;
                var formula = color_formulas[i%color_formulas.Count];
                update.SetFormula(shapeid, fill_foregnd, formula);
                i++;
            }

            update.Execute(page);
        }

        public static void GetShapeColors()
        {
            // Demonstrates how to retrieve all the color formatting
            // as formulas and results

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var layout = new VA.Layout.Grid.GridLayout(5, 5, new VA.Drawing.Size(1, 1), master);
            layout.Origin = new VA.Drawing.Point(0, 5);
            layout.CellSpacing = new VA.Drawing.Size(0, 0);
            layout.RowDirection = VA.Layout.Grid.RowDirection.TopToBottom;
            layout.PerformLayout();
            layout.Render(page);
            
            var srcs = new[]
                           {
                               VA.ShapeSheet.SRCConstants.FillForegnd,
                               VA.ShapeSheet.SRCConstants.FillForegndTrans,
                               VA.ShapeSheet.SRCConstants.FillBkgnd,
                               VA.ShapeSheet.SRCConstants.FillBkgndTrans,
                               VA.ShapeSheet.SRCConstants.ShdwForegnd,
                               VA.ShapeSheet.SRCConstants.ShdwForegndTrans,
                               VA.ShapeSheet.SRCConstants.ShdwForegndTrans,
                               VA.ShapeSheet.SRCConstants.ShdwBkgnd,
                               VA.ShapeSheet.SRCConstants.ShdwBkgndTrans,
                               VA.ShapeSheet.SRCConstants.LineColor,
                               VA.ShapeSheet.SRCConstants.LineColorTrans,
                               VA.ShapeSheet.SRCConstants.Char_Color,
                               VA.ShapeSheet.SRCConstants.Char_ColorTrans,
                               VA.ShapeSheet.SRCConstants.TextBkgnd,
                               VA.ShapeSheet.SRCConstants.TextBkgndTrans
                           };

            var query = new VA.ShapeSheet.Query.CellQuery();
            foreach (var src in srcs)
            {
                query.AddColumn(src);
            }


            var int_shapeids = layout.Nodes.Select( n => (int) n.ShapeID).ToList();
            var results = query.GetResults<double>(page, int_shapeids);
            var formulas = query.GetFormulas(page, int_shapeids);
            var f_and_r = query.GetFormulasAndResults<double>(page, int_shapeids);
        }
    }
}