using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextUtil_Tests : VisioAutomationTest
    {
        public bool Match(string pat, string text)
        {
            var regex = VA.TextUtil.GetRegexForWildcardPattern(pat,true);
            return regex.IsMatch(text);
        }

        [TestMethod]
        public void Text_Case1()
        {
            Assert.IsTrue( Match("*","") );
            Assert.IsTrue(Match("*", "AbC"));

            Assert.IsTrue(Match("A*", "Abc"));
            Assert.IsFalse(Match("A*", "bcA"));

            Assert.IsTrue(Match("*C", "Abc"));
            Assert.IsFalse(Match("*C", "bcA"));

            Assert.IsTrue(Match("A*C", "AbC"));
            Assert.IsFalse(Match("A*C", "AbA"));

            Assert.IsTrue(Match("A*B*C", "A---b---C"));
            Assert.IsFalse(Match("A*B*C", "A---b---A"));

            Assert.IsTrue(Match("A*B?C", "A---bXC"));

        }
    }
}