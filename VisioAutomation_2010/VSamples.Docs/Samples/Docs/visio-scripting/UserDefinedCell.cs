using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/user-defined-cell.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingUserDefinedCellSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape)
        {
            // <sample id="visio-scripting/user-defined-cell#01" doc="visio-scripting/user-defined-cell.md" line="26">
            var udc = new VisioAutomation.Shapes.UserDefinedCellCells();
            udc.Formula = "\"42\"";          // formula, hence escaped quotes
            udc.Prompt = "\"Answer\"";

            // Set user-defined cell "answer" on every selected shape
            client.UserDefinedCell.SetUserDefinedCell(VisioScripting.TargetShapes.Auto, "answer", udc);

            // Read every cell back as resolved values
            var dict = client.UserDefinedCell.GetUserDefinedCellsAsShapeDictionary(
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
