using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class MastersExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void EnumerateFonts()
        {
            var doc1 = this.GetNewDoc();
            var app = doc1.Application;
            var docs = app.Documents;

            var stencil = docs.OpenStencil("basic_u.vss");

            var masters = stencil.Masters;

            var actual = masters.AsEnumerable().ToList();
            for (int i = 0; i < masters.Count; i++)
            {
                Assert.AreEqual( masters[i+1].NameU, actual[i].NameU);
            }
        }
    }
}