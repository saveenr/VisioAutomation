using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingCustomPropTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_CustomProps_Scenarios()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = ss.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = ss.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = ss.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.None();
            ss.Selection.Select(s1);
            ss.Selection.Select(s2);
            ss.Selection.Select(s3);

            var prop_dic0 = ss.CustomProp.Get(null);
            Assert.AreEqual(3, prop_dic0.Count);
            Assert.AreEqual(0, prop_dic0[s1].Count);
            Assert.AreEqual(0, prop_dic0[s2].Count);
            Assert.AreEqual(0, prop_dic0[s3].Count);

            var cp = new CustomPropertyCells();
            cp.Value = "BAR";
            ss.CustomProp.Set(null,"FOO",cp);

            var prop_dic1 = ss.CustomProp.Get(null);
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

            var hasprops0 = ss.CustomProp.Contains(null,"FOO");
            Assert.IsTrue(hasprops0.All(v => v == true));

            ss.CustomProp.Delete(null,"FOO");

            var prop_dic2 = ss.CustomProp.Get(null);
            Assert.AreEqual(3, prop_dic2.Count);
            Assert.AreEqual(0, prop_dic2[s1].Count);
            Assert.AreEqual(0, prop_dic2[s2].Count);
            Assert.AreEqual(0, prop_dic2[s3].Count);

            var hasprops1 = ss.CustomProp.Contains(null,"FOO");
            Assert.IsTrue(hasprops1.All(v => v == false));

            ss.Document.Close(true);
        }
    }
}