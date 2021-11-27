using VA=VisioAutomation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace TestVisioAutomation
{
    
    
    [TestClass()]
    public class FormatStringParserTest
    {

        [TestMethod]
        public void FormatStringParser1()
        {
            var parser1 = new VA.Text.("");
            Assert.AreEqual(0, parser1.Segments.Count);

            var parser2 = new Isotope.Text.FormatStringParser("{0}");
            Assert.AreEqual(1, parser2.Segments.Count);
            Assert.AreEqual(0, parser2.Segments[0].Index);

            var parser3 = new Isotope.Text.FormatStringParser("{0} {2} {0}");
            Assert.AreEqual(3, parser3.Segments.Count);
            Assert.AreEqual(0, parser3.Segments[0].Index);
            Assert.AreEqual(2, parser3.Segments[1].Index);
            Assert.AreEqual(0, parser3.Segments[2].Index);
        }
    }
}
