using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class Scripting_ShapeText_Tests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Scripting_Shape_Text_Set()
        {
            var page1 = this.GetNewPage();
            var stencil = "basic_u.vss";

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var app = page1.Application;
            var documents = app.Documents;
            var stencil_doc = documents.OpenEx(stencil, flags);

            var masters1 = stencil_doc.Masters;
            var masters = new[] { masters1["Rounded Rectangle"], masters1["Ellipse"] };
            var point0 = new VA.Core.Point(1, 2);
            var point1 = new VA.Core.Point(3, 4);
            var points = new[] { point0, point1 };
            MUT.Assert.AreEqual(0, page1.Shapes.Count);

            var shapeids = page1.DropManyU(masters, points);
            MUT.Assert.AreEqual(2, page1.Shapes.Count);
            MUT.Assert.AreEqual(2, shapeids.Length);

            var shapes = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page1.Shapes,shapeids);
            var client = this.GetScriptingClient();
            var names = new[] { "TestName", "TestName2" };
            var texts = names.ToArray();

            var targetshapes = new VisioScripting.TargetShapes(shapes);
            client.Text.SetShapeText(targetshapes, texts);
            client.ShapeSheet.SetShapeName(targetshapes, names);

            for (int i = 0; i < page1.Shapes.Count; i++)
            {
                var shape = shapes[i];
                var name = names[i];
                var text = texts[i];
                MUT.Assert.AreEqual(name, shape.Name);
                MUT.Assert.AreEqual(text, shape.Text);
            }

            page1.Delete(0);
        }
    }
}
