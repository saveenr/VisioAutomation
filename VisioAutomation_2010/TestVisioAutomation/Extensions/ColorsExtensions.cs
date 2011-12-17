using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class ColorsExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void TestAsEnumerable()
        {
            var doc1 = GetNewDoc();
            var colors = doc1.Colors;
            var actual = colors.AsEnumerable().ToList();
            for (int i = 0; i < colors.Count; i++)
            {
                var ex = colors[i];
                var ac = actual[i];
                Assert.AreEqual( ex.Red, ac.Red);
                Assert.AreEqual(ex.Green, ac.Green);
                Assert.AreEqual(ex.Blue, ac.Blue);
            }
        }
    }
}