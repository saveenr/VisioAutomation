using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace TestVisioAutomation
{
    [TestClass]
    public class DocumentsExtensions_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void TestAsEnumerable()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc1 = documents.Add(string.Empty);
            var doc2 = documents.Add(string.Empty);
            var doc3 = documents.Add(string.Empty);

            doc1.Title = "D1";
            doc2.Title = "D2";
            doc3.Title = "D3";

            var actual = documents.AsEnumerable().ToList();
            for (int i = 0; i < documents.Count; i++)
            {
                Assert.AreEqual( documents[i+1].Title , actual[i].Title);
            }
            
            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }

    }
}