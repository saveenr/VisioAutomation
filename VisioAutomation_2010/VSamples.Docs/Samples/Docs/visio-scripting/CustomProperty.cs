using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/custom-property.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingCustomPropertySamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape)
        {
            // <sample id="visio-scripting/custom-property#01" doc="visio-scripting/custom-property.md" line="26">
            var cp = new VisioAutomation.Shapes.CustomPropertyCells();
            cp.Formula = "\"Hello\"";       // values are formulas, hence the escaped quotes
            cp.Label = "\"Greeting\"";
            cp.Type = "0";                 // string

            // Set property "greeting" on every selected shape
            client.CustomProperty.SetCustomProperty(VisioScripting.TargetShapes.Auto, "greeting", cp);

            // Read every property back as resolved values
            var dict = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(
                VisioScripting.TargetShapes.Auto,
                VisioAutomation.Core.CellValueType.Result);

            foreach (var kv in dict)
            {
                foreach (var entry in kv.Value)
                {
                    System.Console.WriteLine("{0}.{1} = {2}", kv.Key.Name, entry.Key, entry.Value.Formula);
                }
            }
            // </sample>
        }
    }
}
