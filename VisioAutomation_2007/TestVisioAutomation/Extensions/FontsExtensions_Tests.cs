using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class FontsExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void EnumerateFonts()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;
            var fonts = doc1.Fonts;

            var expects = fonts.Cast<IVisio.Font>().ToList();
            var actual = fonts.AsEnumerable().ToList();

            Assert.AreEqual(expects.Count,actual.Count);
            for (int i = 0; i < fonts.Count; i++)
            {
                Assert.AreEqual(fonts[i + 1].Name, actual[i].Name);
            }

            page1.Delete(0);
        }
    }
}