using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingCustomPropTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_ShapeSheet_SetNoShapes()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();

            var shapes = new VisioAutomation.Scripting.TargetShapes(s1,s2,s3);
            var shape_ids = shapes.ToShapeIDs();
            var page = client.Page.Get();
            var writer = client.ShapeSheet.GetWriter(page);

            foreach (var shape_id in shape_ids.ShapeIDs)
            {
                writer.SetFormula( (short) shape_id, VA.ShapeSheet.SRCConstants.PinX, "1.0");
            }

            writer.Commit();
            
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_CustomProps_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioAutomation.Scripting.TargetShapes();
            var prop_dic0 = client.CustomProp.Get(targets);
            Assert.AreEqual(3, prop_dic0.Count);
            Assert.AreEqual(0, prop_dic0[s1].Count);
            Assert.AreEqual(0, prop_dic0[s2].Count);
            Assert.AreEqual(0, prop_dic0[s3].Count);

            var cp = new VACUSTPROP.CustomPropertyCells();
            cp.Value = "BAR";
            client.CustomProp.Set(targets, "FOO",cp);

            var prop_dic1 = client.CustomProp.Get(targets);
            Assert.AreEqual(3, prop_dic1.Count);
            Assert.AreEqual(1, prop_dic1[s1].Count);
            Assert.AreEqual(1, prop_dic1[s2].Count);
            Assert.AreEqual(1, prop_dic1[s3].Count);

            var cp1 = prop_dic1[s1]["FOO"];
            var cp2 = prop_dic1[s2]["FOO"];
            var cp3 = prop_dic1[s3]["FOO"];
            Assert.AreEqual("\"BAR\"", cp1.Value.Formula);
            Assert.AreEqual("\"BAR\"", cp2.Value.Formula);
            Assert.AreEqual("\"BAR\"", cp3.Value.Formula);
            

            var hasprops0 = client.CustomProp.Contains(targets,"FOO");
            Assert.IsTrue(hasprops0.All(v => v == true));

            client.CustomProp.Delete(targets,"FOO");

            var prop_dic2 = client.CustomProp.Get(targets);
            Assert.AreEqual(3, prop_dic2.Count);
            Assert.AreEqual(0, prop_dic2[s1].Count);
            Assert.AreEqual(0, prop_dic2[s2].Count);
            Assert.AreEqual(0, prop_dic2[s3].Count);

            var hasprops1 = client.CustomProp.Contains(targets,"FOO");
            Assert.IsTrue(hasprops1.All(v => v == false));

            client.Document.Close(true);
        }
    }
}