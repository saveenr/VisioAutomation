using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core.Text
{
    [MUT.TestClass]
    public class TextUtilTests : Framework.VTest
    {
        public bool Match(string pat, string text)
        {
            var regex = VisioScripting.Helpers.WildcardHelper.GetRegexForWildcardPattern(pat,true);
            return regex.IsMatch(text);
        }

        [MUT.TestMethod]
        public void Text_Case1()
        {
            MUT.Assert.IsTrue(this.Match("*","") );
            MUT.Assert.IsTrue(this.Match("*", "AbC"));
            MUT.Assert.IsTrue(this.Match("A*", "Abc"));
            MUT.Assert.IsTrue(this.Match("*C", "Abc"));
            MUT.Assert.IsFalse(this.Match("A*", "bcA"));
            MUT.Assert.IsFalse(this.Match("*C", "bcA"));
            MUT.Assert.IsTrue(this.Match("A*C", "AbC"));
            MUT.Assert.IsFalse(this.Match("A*C", "AbA"));
            MUT.Assert.IsTrue(this.Match("A*B*C", "A---b---C"));
            MUT.Assert.IsFalse(this.Match("A*B*C", "A---b---A"));
            MUT.Assert.IsTrue(this.Match("A*B?C", "A---bXC"));
        }
    }
}