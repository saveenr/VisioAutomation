using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page custom-properties.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class CustomPropertiesSamples
    {
        public static void Sample_01(IVisio.Document doc)
        {
            // <sample id="custom-properties#01" doc="custom-properties.md" line="9">
            var page = doc.Pages.Add();
            var s1 = page.DrawRectangle(0, 0, 1, 1);

            var cp = new VisioAutomation.Shapes.CustomPropertyCells();
            cp.SetString("Hello World");
            cp.Label = "\"Greeting\"";

            VisioAutomation.Shapes.CustomPropertyHelper.Set(s1, "Greeting", cp);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape s1)
        {
            // <sample id="custom-properties#02" doc="custom-properties.md" line="64">
            var props = VisioAutomation.Shapes.CustomPropertyHelper.GetDictionary(
                            s1,
                            VisioAutomation.Core.CellValueType.Formula);

            foreach (var kv in props)
            {
                System.Console.WriteLine("{0} = {1}", kv.Key, kv.Value.Formula.Value);
            }
            // </sample>
        }

        public static void Sample_03(IVisio.Page page, IVisio.Shape shape, IVisio.Shape s1, IVisio.Shape s2, double a)
        {
            // <sample id="custom-properties#03" doc="custom-properties.md" line="81">
            var pairs = VisioAutomation.Core.ShapeIDPairs.FromShapes(s1, s2);
            var per_shape = VisioAutomation.Shapes.CustomPropertyHelper.GetDictionary(
                                page,
                                pairs,
                                VisioAutomation.Core.CellValueType.Formula);
            // per_shape is a List<CustomPropertyDictionary>, one entry per shape
            // </sample>
        }

        public static void Sample_04(IVisio.Shape s1)
        {
            // <sample id="custom-properties#04" doc="custom-properties.md" line="92">
            int count = VisioAutomation.Shapes.CustomPropertyHelper.GetCount(s1);

            System.Collections.Generic.List<string> names =
                VisioAutomation.Shapes.CustomPropertyHelper.GetNames(s1);
            // </sample>
        }

        public static void Sample_05(IVisio.Shape s1)
        {
            // <sample id="custom-properties#05" doc="custom-properties.md" line="101">
            bool exists = VisioAutomation.Shapes.CustomPropertyHelper.Contains(s1, "Greeting");
            // </sample>
        }

        public static void Sample_06(IVisio.Shape s1)
        {
            // <sample id="custom-properties#06" doc="custom-properties.md" line="107">
            VisioAutomation.Shapes.CustomPropertyHelper.Delete(s1, "Greeting");
            // </sample>
        }

        public static void Sample_07()
        {
            // <sample id="custom-properties#07" doc="custom-properties.md" line="113">
            bool ok = VisioAutomation.Shapes.CustomPropertyHelper.IsValidName("MyProp");
            // </sample>
        }
    }
}
