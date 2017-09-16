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

            // Set some properties on it
            CustomPropertyHelper.Set(s1, "FOO1", "BAR1");
            CustomPropertyHelper.Set(s1, "FOO2", "BAR2");
            CustomPropertyHelper.Set(s1, "FOO3", "BAR3");

            // Delete one of those properties
            CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            CustomPropertyHelper.Set(s1, "FOO3", "BAR3updated");

            // retrieve all the properties
            var props = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);

            var cp_foo1 = props["FOO1"];
            var cp_foo2 = props["FOO2"];
            var cp_foo3 = props["FOO3"];
        }
    }
}