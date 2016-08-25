using VisioAutomation.Utilities;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;

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
            VACUSTPROP.CustomPropertyHelper.Set(s1, "FOO1", "BAR1");
            VACUSTPROP.CustomPropertyHelper.Set(s1, "FOO2", "BAR2");
            VACUSTPROP.CustomPropertyHelper.Set(s1, "FOO3", "BAR3");

            // Delete one of those properties
            VACUSTPROP.CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            string formula = Convert.StringToFormulaString("BAR3updated");
            VACUSTPROP.CustomPropertyHelper.Set(s1, "FOO3", formula);

            // retrieve all the properties
            var props = VACUSTPROP.CustomPropertyHelper.Get(s1);
        }
    }
}