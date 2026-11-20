using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shapesheet/modify-the-shapesheet.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapesheetModifyTheShapesheetSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="shapesheet/modify-the-shapesheet#01" doc="shapesheet/modify-the-shapesheet.md" line="11">
            var pinx_cell = shape.CellsU["PinX"];
            pinx_cell.FormulaU = "5.0";
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape1)
        {
            // <sample id="shapesheet/modify-the-shapesheet#02" doc="shapesheet/modify-the-shapesheet.md" line="22">
            var writer = new VA.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.Core.SrcConstants.XFormPinX, "5.0");
            writer.Commit(shape1, VA.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape1)
        {
            // <sample id="shapesheet/modify-the-shapesheet#03" doc="shapesheet/modify-the-shapesheet.md" line="32">
            var writer = new VA.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.Core.SrcConstants.XFormPinX, 5.0);
            writer.Commit(shape1, VA.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape1)
        {
            // <sample id="shapesheet/modify-the-shapesheet#04" doc="shapesheet/modify-the-shapesheet.md" line="40">
            var writer = new VA.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.Core.SrcConstants.XFormPinX, 5.0);
            writer.SetValue(VA.Core.SrcConstants.XFormPinY, 3.5);
            writer.Commit(shape1, VA.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_05(IVisio.Page page1, VA.ShapeSheet.Writers.SrcWriter writer, IVisio.Master master1)
        {
            // <sample id="shapesheet/modify-the-shapesheet#05" doc="shapesheet/modify-the-shapesheet.md" line="49">
            writer.Commit(master1, VA.Core.CellValueType.Formula);
            writer.Commit(page1, VA.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_06(IVisio.Shape shape1)
        {
            // <sample id="shapesheet/modify-the-shapesheet#06" doc="shapesheet/modify-the-shapesheet.md" line="58">
            var writer = new VA.ShapeSheet.Writers.SrcWriter();
            writer.SetValue(VA.Core.SrcConstants.XFormPinX, 5.0);
            writer.Commit(shape1, VA.Core.CellValueType.Result);
            // </sample>
        }

        public static void Sample_07(IVisio.Shape shape1, IVisio.Shape shape2, IVisio.Shape shape3, IVisio.Page page1, VisioAutomation.Core.Src src_pinx, VisioAutomation.Core.Src src_piny)
        {
            // <sample id="shapesheet/modify-the-shapesheet#07" doc="shapesheet/modify-the-shapesheet.md" line="70">
            var writer = new VA.ShapeSheet.Writers.SidSrcWriter();
            writer.SetValue(shape1.ID16, src_pinx, 0.5);
            writer.SetValue(shape1.ID16, src_piny, 0.5);
            writer.SetValue(shape2.ID16, src_pinx, 1.5);
            writer.SetValue(shape2.ID16, src_piny, 1.5);
            writer.SetValue(shape3.ID16, src_pinx, 2.5);
            writer.SetValue(shape3.ID16, src_piny, 2.5);
            writer.Commit(page1, VA.Core.CellValueType.Formula);
            // </sample>
        }
    }
}
