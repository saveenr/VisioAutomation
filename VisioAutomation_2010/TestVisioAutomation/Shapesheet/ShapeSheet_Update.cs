using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeSheet_Update : VisioAutomationTest
    {
        private static VA.ShapeSheet.SRC src_fg = VA.ShapeSheet.SRCConstants.FillForegnd;
        private static VA.ShapeSheet.SRC src_bg = VA.ShapeSheet.SRCConstants.FillBkgnd;
        private static VA.ShapeSheet.SRC src_fillpat = VA.ShapeSheet.SRCConstants.FillPattern;
        private static VA.ShapeSheet.SRC src_pinx = VA.ShapeSheet.SRCConstants.PinX;
        private static VA.ShapeSheet.SRC src_piny = VA.ShapeSheet.SRCConstants.PinY;
        private static VA.ShapeSheet.SRC src_linepat = VA.ShapeSheet.SRCConstants.LinePattern;

        [TestMethod]
        public void Set_Cell_On_Shape()
        {
            var page1 = GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            string fg_formula = VA.Convert.ColorToFormulaRGB(255, 0, 0);
            string bg_formula = VA.Convert.ColorToFormulaRGB(255, 128, 0);

            // Setup the modifications to the cell values
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            update.SetFormula(src_fg, fg_formula);
            update.SetFormula(src_bg, bg_formula);
            update.SetFormula(src_fillpat, 40);
            update.SetResult(src_linepat, 7, IVisio.VisUnitCodes.visNoCast);
            update.Execute(shape1);

            // Build the query
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.AddColumn(src_fg);
            var col_bg = query.AddColumn(src_bg);
            var col_fillpat = query.AddColumn(src_fillpat);
            var col_linepat = query.AddColumn(src_linepat);

            // Retrieve the values
            var formulas = query.GetFormulas(shape1);
            var results = query.GetResults<double>(shape1);

            // Verify
            Assert.AreEqual("RGB(255,0,0)", formulas[0, col_fg]);
            Assert.AreEqual("RGB(255,128,0)", formulas[0, col_bg]);
            Assert.AreEqual("40", formulas[0, col_fillpat]);
            Assert.AreEqual(7.0, results[0, col_linepat]);

            page1.Delete(0);
        }

        [TestMethod]
        public void Set_Cells_Formulas_On_Shapes()
        {
            var page1 = GetNewPage();

            var shape1 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape2 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape3 = page1.DrawRectangle(-1, -1, 0, 0);


            // Set the formulas
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            update.SetFormula(shape1.ID16, src_pinx, 0.5);
            update.SetFormula(shape1.ID16, src_piny, 0.5);
            update.SetFormula(shape2.ID16, src_pinx, 1.5);
            update.SetFormula(shape2.ID16, src_piny, 1.5);
            update.SetFormula(shape3.ID16, src_pinx, 2.5);
            update.SetFormula(shape3.ID16, src_piny, 2.5);
            update.Execute(page1);

            // Verify that the formulas were set
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_pinx = query.AddColumn(src_pinx);
            var col_piny = query.AddColumn(src_piny);

            var shapeids = new int[] {shape1.ID, shape2.ID, shape3.ID};

            var r = query.GetFormulasAndResults<double>(page1, shapeids);

            Assert.AreEqual("0.5 in", r[0, col_pinx].Formula);
            Assert.AreEqual("0.5 in", r[0, col_piny].Formula);
            Assert.AreEqual("1.5 in", r[1, col_pinx].Formula);
            Assert.AreEqual("1.5 in", r[1, col_piny].Formula);
            Assert.AreEqual("2.5 in", r[2, col_pinx].Formula);
            Assert.AreEqual("2.5 in", r[2, col_piny].Formula);

            Assert.AreEqual(0.5, r[0, col_pinx].Result);
            Assert.AreEqual(0.5, r[0, col_piny].Result);
            Assert.AreEqual(1.5, r[1, col_pinx].Result);
            Assert.AreEqual(1.5, r[1, col_piny].Result);
            Assert.AreEqual(2.5, r[2, col_pinx].Result);
            Assert.AreEqual(2.5, r[2, col_piny].Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void Set_1_Cell_on_Many_Shapes()
        {
            var page1 = GetNewPage();

            var unitcode_nocast = IVisio.VisUnitCodes.visNoCast;
            var src_pinx = VA.ShapeSheet.SRCConstants.PinX;

            // draw a simple shape
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 1, 2, 2);
            var s3 = page1.DrawRectangle(2, 2, 3, 3);

            // format it with setformulas
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            update.SetResult((short)s1.ID, src_pinx, 5.0, unitcode_nocast);
            update.SetResult((short)s2.ID, src_pinx, 6.0, unitcode_nocast);
            update.SetResult((short)s3.ID, src_pinx, 7.0, unitcode_nocast);

            update.Execute(page1);

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_pinx = query.AddColumn(src_pinx);
            var shapeids = new[] { s1.ID, s2.ID, s3.ID };

            var results = query.GetResults<double>(page1, shapeids);
            Assert.AreEqual(5.0, results[0, col_pinx]);
            Assert.AreEqual(6.0, results[1, col_pinx]);
            Assert.AreEqual(7.0, results[2, col_pinx]);

            page1.Delete(0);
        }


    }
}