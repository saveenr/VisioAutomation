using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VACUSTOMPROP = VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;

namespace TestVisioAutomation
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

            client.Selection.None();

            client.ShapeSheet.SetFormula(null, new [] {VA.ShapeSheet.SRCConstants.PinX}, new []{"1.0"}, 0 );
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

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var prop_dic0 = client.CustomProp.Get(null);
            Assert.AreEqual(3, prop_dic0.Count);
            Assert.AreEqual(0, prop_dic0[s1].Count);
            Assert.AreEqual(0, prop_dic0[s2].Count);
            Assert.AreEqual(0, prop_dic0[s3].Count);

            var cp = new VACUSTOMPROP.CustomPropertyCells();
            cp.Value = "BAR";
            client.CustomProp.Set(null,"FOO",cp);

            var prop_dic1 = client.CustomProp.Get(null);
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

            var hasprops0 = client.CustomProp.Contains(null,"FOO");
            Assert.IsTrue(hasprops0.All(v => v == true));

            client.CustomProp.Delete(null,"FOO");

            var prop_dic2 = client.CustomProp.Get(null);
            Assert.AreEqual(3, prop_dic2.Count);
            Assert.AreEqual(0, prop_dic2[s1].Count);
            Assert.AreEqual(0, prop_dic2[s2].Count);
            Assert.AreEqual(0, prop_dic2[s3].Count);

            var hasprops1 = client.CustomProp.Contains(null,"FOO");
            Assert.IsTrue(hasprops1.All(v => v == false));

            client.Document.Close(true);
        }
    }
}