using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;

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
            VA.CustomProperties.CustomPropertyHelper.Set(s1, "FOO1", "BAR1");
            VA.CustomProperties.CustomPropertyHelper.Set(s1, "FOO2", "BAR2");
            VA.CustomProperties.CustomPropertyHelper.Set(s1, "FOO3", "BAR3");

            // Delete one of those properties
            VA.CustomProperties.CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            string formula = VA.Convert.StringToFormulaString("BAR3updated");
            VA.CustomProperties.CustomPropertyHelper.Update(s1, "FOO3", formula);

            // retrieve all the properties
            var props = VA.CustomProperties.CustomPropertyHelper.Get(s1);
        }
    }
}