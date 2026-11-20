using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/lock.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingLockSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Shape shape)
        {
            // <sample id="visio-scripting/lock#01" doc="visio-scripting/lock.md" line="16">
            var locks = new VisioAutomation.Shapes.LockCells();
            locks.MoveX = 1;     // disallow horizontal move
            locks.MoveY = 1;     // disallow vertical move
            locks.Width = 1;     // disallow resize-width
            locks.Height = 1;    // disallow resize-height

            // Pin every selected shape in place
            client.Lock.SetLockCells(VisioScripting.TargetShapes.Auto, locks);

            // Read them back as resolved values
            var dict = client.Lock.GetLockCells(VisioScripting.TargetShapes.Auto,
                                                  VisioAutomation.Core.CellValueType.Result);
            foreach (var kv in dict)
            {
                System.Console.WriteLine("Shape #{0}: MoveX={1} MoveY={2}",
                    kv.Key, kv.Value.MoveX, kv.Value.MoveY);
            }
            // </sample>
        }
    }
}
