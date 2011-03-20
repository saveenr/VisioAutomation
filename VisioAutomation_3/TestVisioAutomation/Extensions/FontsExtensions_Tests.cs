using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VA=VisioAutomation;

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

            var fonts = doc1.Fonts.AsEnumerable().ToList();
            Assert.IsTrue(fonts.Count > 0);

            page1.Delete(0);
        }

        [TestMethod]
        public void FindFontByName()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            var f1 = VA.Text.TextHelper.FindFontWithName(doc1.Fonts, "Arial");
            Assert.IsNotNull(f1);
            var f2 = VA.Text.TextHelper.FindFontWithName(doc1.Fonts, "aRIAL");
            Assert.IsNotNull(f2);
            Assert.AreSame(f1, f2);

            var f3 = VA.Text.TextHelper.FindFontWithName(doc1.Fonts, "UnknownFont123456");
            Assert.IsNull(f3);

            page1.Delete(0);
        }
    }
}