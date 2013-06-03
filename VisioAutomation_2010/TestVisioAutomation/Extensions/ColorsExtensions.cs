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
                var expected_color = colors[i];
                var actual_color = actual[i];
                Assert.AreEqual(expected_color.Red, actual_color.Red);
                Assert.AreEqual(expected_color.Green, actual_color.Green);
                Assert.AreEqual(expected_color.Blue, actual_color.Blue);
            }
            doc1.Close(true);
        }
    }
}