using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class LayersExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void TestAsEnumerable()
        {
            var doc1 = GetNewDoc();
            var page1 = doc1.Pages[1];

            var layers = page1.Layers;
            layers.Add("FOO");
            layers.Add("BAR");
            layers.Add("BEER");

            var actual = layers.AsEnumerable().ToList();
            for (int i = 0; i < layers.Count; i++)
            {
                var ex = layers[i+1];
                var ac = actual[i];
                Assert.AreEqual(ex.NameU, ac.NameU);
            }
        }
    }
}