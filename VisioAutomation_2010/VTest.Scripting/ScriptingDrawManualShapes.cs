using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Models
{
    public class ScriptingDrawManualShapes : Framework.VTest
    {
        [MUT.TestMethod]
        public void Scripting_Draw_RectangleLineOval_0()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var pagesize = new VA.Core.Size(4, 4);

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var shape_rect = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 3, 3);
            var shape_line = client.Draw.DrawLine(VisioScripting.TargetPage.Auto, 0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.DrawOval(VisioScripting.TargetPage.Auto, 0.2, 1, 3.8, 2);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Draw_BezierPolyLine_0()
        {
            var points = new[]
                {
                    new VA.Core.Point(0, 0),
                    new VA.Core.Point(2, 0.5),
                    new VA.Core.Point(2, 2),
                    new VA.Core.Point(3, 0.5)
                };
            var pagesize = new VA.Core.Size(4, 4);

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Draw the Shapes
            var shape_bezier = client.Draw.DrawBezier(VisioScripting.TargetPage.Auto, points);
            var shape_polyline = client.Draw.DrawPolyLine(VisioScripting.TargetPage.Auto, points);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

    }
}