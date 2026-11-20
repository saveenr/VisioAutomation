using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page user-defined-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class UserDefinedCellsSamples
    {
        public static void Sample_01(IVisio.Shape shape1)
        {
            // <sample id="user-defined-cells#01" doc="user-defined-cells.md" line="9">
            var udcells = VisioAutomation.Shapes.UserDefinedCellHelper.GetDictionary(
                              shape1,
                              VisioAutomation.Core.CellValueType.Formula);

            foreach (var kv in udcells)
            {
                System.Console.WriteLine("{0} = {1}", kv.Key, kv.Value.Formula);
            }
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape, IVisio.Shape shape1, IVisio.Shape shape2, IVisio.Shape shape3, IVisio.Page page1, double a)
        {
            // <sample id="user-defined-cells#02" doc="user-defined-cells.md" line="22">
            var pairs = VisioAutomation.Core.ShapeIDPairs.FromShapes(shape1, shape2, shape3);
            var per_shape = VisioAutomation.Shapes.UserDefinedCellHelper.GetDictionary(
                                page1,
                                pairs,
                                VisioAutomation.Core.CellValueType.Formula);
            // per_shape is a List<UserDefinedCellDictionary>, one entry per shape
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape1)
        {
            // <sample id="user-defined-cells#03" doc="user-defined-cells.md" line="35">
            var cells = new VisioAutomation.Shapes.UserDefinedCellCells();
            cells.SetString("the answer");      // string literal
            cells.Prompt = "\"Hint shown to the user\"";

            VisioAutomation.Shapes.UserDefinedCellHelper.Set(shape1, "MyCell", cells);
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape1)
        {
            // <sample id="user-defined-cells#04" doc="user-defined-cells.md" line="61">
            int count = VisioAutomation.Shapes.UserDefinedCellHelper.GetCount(shape1);

            System.Collections.Generic.List<string> names =
                VisioAutomation.Shapes.UserDefinedCellHelper.GetNames(shape1);
            // </sample>
        }

        public static void Sample_05(IVisio.Shape shape1)
        {
            // <sample id="user-defined-cells#05" doc="user-defined-cells.md" line="70">
            bool exists = VisioAutomation.Shapes.UserDefinedCellHelper.Contains(shape1, "MyCell");
            // </sample>
        }

        public static void Sample_06(IVisio.Shape shape1)
        {
            // <sample id="user-defined-cells#06" doc="user-defined-cells.md" line="76">
            VisioAutomation.Shapes.UserDefinedCellHelper.Delete(shape1, "MyCell");
            // </sample>
        }

        public static void Sample_07()
        {
            // <sample id="user-defined-cells#07" doc="user-defined-cells.md" line="82">
            bool ok = VisioAutomation.Shapes.UserDefinedCellHelper.IsValidName("MyCell");
            // or, throw on invalid:
            VisioAutomation.Shapes.UserDefinedCellHelper.CheckValidName("MyCell");
            // </sample>
        }
    }
}
