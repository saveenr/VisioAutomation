using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class WindowsExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void TestAsEnumerable()
        {
            var doc1 = GetNewDoc();
            var doc2 = GetNewDoc();
            var app = doc1.Application;
            var windows = app.Windows;
            var actual = windows.AsEnumerable().ToList();
            for (int i = 0; i < windows.Count; i++)
            {
                var ex = windows[(short)(i+1)];
                var ac = actual[i];
                Assert.AreEqual( ex.ID, ac.ID);
            }
        }
    }
}