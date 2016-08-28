using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class ShapeSheetWriterTests : VisioAutomationTest
    {
        private static readonly VA.ShapeSheet.SRC src_pinx = VA.ShapeSheet.SRCConstants.PinX;
        private static readonly VA.ShapeSheet.SRC src_piny = VA.ShapeSheet.SRCConstants.PinY;
        private static readonly VA.ShapeSheet.SRC src_linepat = VA.ShapeSheet.SRCConstants.LinePattern;

        [TestMethod]
        public void ShapeSheet_Writer_Formulas_MultipleShapes()
        {
            var page1 = this.GetNewPage();

            var shape1 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape2 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape3 = page1.DrawRectangle(-1, -1, 0, 0);


            // Set the formulas
            var writer = new FormulaWriterSIDSRC();
            writer.SetFormula(shape1.ID16, ShapeSheetWriterTests.src_pinx, 0.5);
            writer.SetFormula(shape1.ID16, ShapeSheetWriterTests.src_piny, 0.5);
            writer.SetFormula(shape2.ID16, ShapeSheetWriterTests.src_pinx, 1.5);
            writer.SetFormula(shape2.ID16, ShapeSheetWriterTests.src_piny, 1.5);
            writer.SetFormula(shape3.ID16, ShapeSheetWriterTests.src_pinx, 2.5);
            writer.SetFormula(shape3.ID16, ShapeSheetWriterTests.src_piny, 2.5);
            writer.Commit(page1);

            // Verify that the formulas were set
            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_pinx = query.AddCell(ShapeSheetWriterTests.src_pinx, "PinX");
            var col_piny = query.AddCell(ShapeSheetWriterTests.src_piny, "PinY");

            var shapeids = new[] { shape1.ID, shape2.ID, shape3.ID };

            var ss1 = new ShapeSheetSurface(page1);
            var rf = query.GetFormulas(ss1, shapeids);
            var rr = query.GetResults<double>(ss1, shapeids);

            AssertUtil.AreEqual("0.5 in", 0.5, rf[0].Cells[col_pinx], rr[0].Cells[col_pinx]);
            AssertUtil.AreEqual("0.5 in", 0.5, rf[0].Cells[col_piny], rr[0].Cells[col_piny]);
            AssertUtil.AreEqual("1.5 in", 1.5, rf[1].Cells[col_pinx], rr[1].Cells[col_pinx]);
            AssertUtil.AreEqual("1.5 in", 1.5, rf[1].Cells[col_piny], rr[1].Cells[col_piny]);
            AssertUtil.AreEqual("2.5 in", 2.5, rf[2].Cells[col_pinx], rr[2].Cells[col_pinx]);
            AssertUtil.AreEqual("2.5 in", 2.5, rf[2].Cells[col_piny], rr[2].Cells[col_piny]);

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ResultsInt_SingleShape()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new ResultWriterSRC();
            writer.SetResult(ShapeSheetWriterTests.src_linepat, 7, IVisio.VisUnitCodes.visNumber);
            writer.Commit(shape1);

            // Build the query
            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_linepat = query.AddCell(ShapeSheetWriterTests.src_linepat,"LinePattern");

            // Retrieve the values
            var ss1 = new ShapeSheetSurface(shape1);
            var data = query.GetFormulasAndResults<double>(ss1);

            // Verify
            AssertUtil.AreEqual("7", 7, data.Cells[col_linepat]);
            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ResultsString_SingleShape()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new ResultWriterSRC();
            writer.SetResult(ShapeSheetWriterTests.src_linepat, "7", IVisio.VisUnitCodes.visNumber);
            writer.Commit(shape1);

            // Build the query
            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_linepat = query.AddCell(ShapeSheetWriterTests.src_linepat, "LinePattern");

            // Retrieve the values
            var ss1 = new ShapeSheetSurface(shape1);
            var data = query.GetFormulasAndResults<double>(ss1);

            // Verify
            AssertUtil.AreEqual("7", 7, data.Cells[col_linepat]);
            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ResultsDouble_MultipleShapes()
        {
            var page1 = this.GetNewPage();

            var shape1 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape2 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape3 = page1.DrawRectangle(-1, -1, 0, 0);


            // Set the formulas
            var writer = new ResultWriterSIDSRC();
            writer.SetResult( new SIDSRC(shape1.ID16, ShapeSheetWriterTests.src_pinx), 0.5, IVisio.VisUnitCodes.visNumber);
            writer.SetResult( new SIDSRC(shape1.ID16, ShapeSheetWriterTests.src_piny), 0.5, IVisio.VisUnitCodes.visNumber);
            writer.SetResult( new SIDSRC(shape2.ID16, ShapeSheetWriterTests.src_pinx), 1.5, IVisio.VisUnitCodes.visNumber);
            writer.SetResult( new SIDSRC(shape2.ID16, ShapeSheetWriterTests.src_piny), 1.5, IVisio.VisUnitCodes.visNumber);
            writer.SetResult( new SIDSRC(shape3.ID16, ShapeSheetWriterTests.src_pinx), 2.5, IVisio.VisUnitCodes.visNumber);
            writer.SetResult( new SIDSRC(shape3.ID16, ShapeSheetWriterTests.src_piny), 2.5, IVisio.VisUnitCodes.visNumber);
            writer.Commit(page1);

            // Verify that the formulas were set
            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_pinx = query.AddCell(ShapeSheetWriterTests.src_pinx,"PinX");
            var col_piny = query.AddCell(ShapeSheetWriterTests.src_piny, "PinY");

            var shapeids = new[] { shape1.ID, shape2.ID, shape3.ID };

            var ss1 = new ShapeSheetSurface(page1);
            var rf = query.GetFormulas(ss1, shapeids);
            var rr = query.GetResults<double>(ss1, shapeids);

            AssertUtil.AreEqual("0.5 in", 0.5, rf[0].Cells[col_pinx], rr[0].Cells[col_pinx]);
            AssertUtil.AreEqual("0.5 in", 0.5, rf[0].Cells[col_piny], rr[0].Cells[col_piny]);
            AssertUtil.AreEqual("1.5 in", 1.5, rf[1].Cells[col_pinx], rr[1].Cells[col_pinx]);
            AssertUtil.AreEqual("1.5 in", 1.5, rf[1].Cells[col_piny], rr[1].Cells[col_piny]);
            AssertUtil.AreEqual("2.5 in", 2.5, rf[2].Cells[col_pinx], rr[2].Cells[col_pinx]);
            AssertUtil.AreEqual("2.5 in", 2.5, rf[2].Cells[col_piny], rr[2].Cells[col_piny]);

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ConsistencyChecking()
        {
            this.Check_Consistent_ResultTypes();
        }
        


        public void Check_Consistent_ResultTypes()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new ResultWriterSRC();
            writer.SetResult(ShapeSheetWriterTests.src_linepat, "7", IVisio.VisUnitCodes.visNumber);
            writer.SetResult(VA.ShapeSheet.SRCConstants.PinX, 2, IVisio.VisUnitCodes.visNumber);
            writer.Commit(shape1);

            // Build the query
            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_linepat = query.AddCell(ShapeSheetWriterTests.src_linepat, "LinePattern");
            var col_pinx = query.AddCell(VA.ShapeSheet.SRCConstants.PinX, "PinX");

            // Retrieve the values
            var ss1 = new ShapeSheetSurface(shape1);
            var data = query.GetFormulasAndResults<double>(ss1);

            // Verify
            AssertUtil.AreEqual("7", 7, data.Cells[col_linepat]);
            AssertUtil.AreEqual("2 in", 2, data.Cells[col_pinx]);
            
            page1.Delete(0);
        }
    }
}