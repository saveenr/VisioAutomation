using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;

namespace VisioAutomationSamples
{
    public static class CustomPropertySamples
    {
        public static void SetCustomProperties()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Draw a shape
            var s1 = page.DrawRectangle(1, 1, 4, 3);

            int cp_type = 0; // string type

            // Set some properties on it
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO2", "\"BAR2\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3\"", cp_type);

            // Delete one of those properties
            CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3updated\"", cp_type);

            // retrieve all the properties
            var props = CustomPropertyHelper.GetCellsAsDictionary(s1, CellValueType.Formula);

            var cp_foo1 = props["FOO1"];
            // var cp_foo2 = props["FOO2"]; // there is not prop called FOO2
            var cp_foo3 = props["FOO3"];
        }
    }
}