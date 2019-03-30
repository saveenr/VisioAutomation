using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet.Query;
using VisioAutomation.ShapeSheet.Writers;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class ShapeSheetWriterTests : VisioAutomationTest
    {
        private static readonly VA.ShapeSheet.Src XFormPinX = VA.ShapeSheet.SrcConstants.XFormPinX;
        private static readonly VA.ShapeSheet.Src XFormPinY = VA.ShapeSheet.SrcConstants.XFormPinY;
        private static readonly VA.ShapeSheet.Src LinePattern = VA.ShapeSheet.SrcConstants.LinePattern;

        [TestMethod]
        public void ShapeSheet_Writer_Formulas_MultipleShapes()
        {
            var page1 = this.GetNewPage();

            var shape1 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape2 = page1.DrawRectangle(-1, -1, 0, 0);
            var shape3 = page1.DrawRectangle(-1, -1, 0, 0);


            // Set the formulas
            var writer = new SidSrcWriter();
            writer.SetValue(shape1.ID16, XFormPinX, 0.5);
            writer.SetValue(shape1.ID16, XFormPinY, 0.5);
            writer.SetValue(shape2.ID16, XFormPinX, 1.5);
            writer.SetValue(shape2.ID16, XFormPinY, 1.5);
            writer.SetValue(shape3.ID16, XFormPinX, 2.5);
            writer.SetValue(shape3.ID16, XFormPinY, 2.5);

            writer.Commit(page1, VA.ShapeSheet.CellValueType.Formula);

            // Verify that the formulas were set
            var query = new CellQuery();
            var col_pinx = query.Columns.Add(XFormPinX, nameof(XFormPinX));
            var col_piny = query.Columns.Add(XFormPinY, nameof(XFormPinY));

            var shapeids = new[] { shape1.ID, shape2.ID, shape3.ID };

            var data_formulas = query.GetFormulas(page1, shapeids);
            var data_results = query.GetResults<double>(page1, shapeids);

            AssertUtil.AreEqual(("0.5 in", 0.5), (data_formulas[0][col_pinx], data_results[0][col_pinx]));
            AssertUtil.AreEqual(("0.5 in", 0.5), (data_formulas[0][col_piny], data_results[0][col_piny]));
            AssertUtil.AreEqual(("1.5 in", 1.5), (data_formulas[1][col_pinx], data_results[1][col_pinx]));
            AssertUtil.AreEqual(("1.5 in", 1.5), (data_formulas[1][col_piny], data_results[1][col_piny]));
            AssertUtil.AreEqual(("2.5 in", 2.5), (data_formulas[2][col_pinx], data_results[2][col_pinx]));
            AssertUtil.AreEqual(("2.5 in", 2.5), (data_formulas[2][col_piny], data_results[2][col_piny]));

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ResultsInt_SingleShape()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new SrcWriter();
            writer.SetValue(LinePattern, 7);

            writer.CommitResults(shape1);

            // Build the query
            var query = new CellQuery();
            var col_linepat = query.Columns.Add(LinePattern,nameof(LinePattern));

            // Retrieve the values
            var data_formulas = query.GetFormulas(shape1);
            var data_results = query.GetResults<double>(shape1);

            int rownum=0;
            // Verify
            Assert.AreEqual("7", data_formulas[rownum][col_linepat]);
            Assert.AreEqual(7, data_results[rownum][col_linepat]);
            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_Write_nothing()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new SrcWriter();
            writer.CommitFormulas(shape1);

            page1.Delete(0);
        }

        [TestMethod]
        public void ShapeSheet_Writer_ResultsString_SingleShape()
        {
            var page1 = this.GetNewPage();
            var shape1 = page1.DrawRectangle(0, 0, 1, 1);

            // Setup the modifications to the cell values
            var writer = new SrcWriter();
            writer.SetValue(LinePattern, "7");
            writer.CommitFormulas(shape1);

            // Build the query
            var query = new CellQuery();
            var col_linepat = query.Columns.Add(LinePattern, nameof(LinePattern));

            // Retrieve the values
            var data_formulas = query.GetFormulas(shape1);
            var data_results = query.GetResults<double>(shape1);

            int rownum = 0;
            // Verify
            Assert.AreEqual("7", data_formulas[rownum][col_linepat]);
            Assert.AreEqual(7, data_results[rownum][col_linepat]);
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
            var writer = new SidSrcWriter();
            writer.SetValue( shape1.ID16, XFormPinX, 0.5);
            writer.SetValue( shape1.ID16, XFormPinY, 0.5);
            writer.SetValue( shape2.ID16, XFormPinX, 1.5);
            writer.SetValue( shape2.ID16, XFormPinY, 1.5);
            writer.SetValue( shape3.ID16, XFormPinX, 2.5);
            writer.SetValue( shape3.ID16, XFormPinY, 2.5);

            writer.Commit(page1, VA.ShapeSheet.CellValueType.Result);

            // Verify that the formulas were set
            var query = new CellQuery();
            var col_pinx = query.Columns.Add(XFormPinX, nameof(XFormPinX));
            var col_piny = query.Columns.Add(XFormPinY, nameof(XFormPinY));

            var shapeids = new[] { shape1.ID, shape2.ID, shape3.ID };

            var data_formulas = query.GetFormulas(page1, shapeids);
            var data_results = query.GetResults<double>(page1, shapeids);

            AssertUtil.AreEqual(("0.5 in", 0.5), (data_formulas[0][col_pinx], data_results[0][col_pinx]));
            AssertUtil.AreEqual(("0.5 in", 0.5), (data_formulas[0][col_piny], data_results[0][col_piny]));
            AssertUtil.AreEqual(("1.5 in", 1.5), (data_formulas[1][col_pinx], data_results[1][col_pinx]));
            AssertUtil.AreEqual(("1.5 in", 1.5), (data_formulas[1][col_piny], data_results[1][col_piny]));
            AssertUtil.AreEqual(("2.5 in", 2.5), (data_formulas[2][col_pinx], data_results[2][col_pinx]));
            AssertUtil.AreEqual(("2.5 in", 2.5), (data_formulas[2][col_piny], data_results[2][col_piny]));

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
            var writer = new SrcWriter();
            writer.SetValue(LinePattern, "7");
            writer.SetValue(XFormPinX, 2);
            writer.CommitResults(shape1);

            // Build the query
            var query = new CellQuery();
            var col_linepat = query.Columns.Add(LinePattern, nameof(LinePattern));
            var col_pinx = query.Columns.Add(XFormPinX, nameof(XFormPinX));

            // Retrieve the values
            var data_formulas = query.GetFormulas(shape1);
            var data_results = query.GetResults<double>(shape1);

            int rownum = 0;
            // Verify
            Assert.AreEqual("7", data_formulas[rownum][col_linepat]);
            Assert.AreEqual(7, data_results[rownum][col_linepat]);

            Assert.AreEqual("2 in", data_formulas[rownum][col_pinx]);
            Assert.AreEqual(2, data_results[rownum][col_pinx]);
            
            page1.Delete(0);
        }
    }
}