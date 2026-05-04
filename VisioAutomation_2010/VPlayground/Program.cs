using IVisio = Microsoft.Office.Interop.Visio;
using VAS = VisioScripting;
using VA = VisioAutomation;

namespace VPlayground
{
    // Throwaway harness for trying out old issues, ad-hoc API exploration, and
    // anything else that doesn't belong in a test or sample. Edit Main() freely
    // -- the project is here for scratch work, not curated content. Don't
    // commit issue-specific changes unless you're deliberately keeping them.
    //
    // The project references VisioAutomation, VisioAutomation.Models, and
    // VisioScripting, so anything in those is reachable directly. Microsoft.Office.Interop.Visio
    // is on the same reference list as the rest of the solution.
    //
    // Visio is left alive at the end of Main() so you can inspect the result.
    // Close Visio manually when you're done (or call app.Quit() at the end).

    public static class Program
    {
        public static void Main(string[] args)
        {
            var app = new IVisio.Application();
            var client = new VAS.Client(app);
            client.Document.NewDocument();
            client.Draw.DrawRectangle(VAS.TargetPage.Auto, new VA.Core.Rectangle(0, 0, 4, 2));
            client.Text.SetShapeText(VAS.TargetShapes.Auto, new[] { "Hello, Visio!" });
        }
    }
}
