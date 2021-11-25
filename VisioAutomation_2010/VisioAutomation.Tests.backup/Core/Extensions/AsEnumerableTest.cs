using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class AsEnumerableTests : VisioAutomationTest
    {
        [TestMethod]
        public void VerifyAsEnu()
        {
            var page1 = GetNewPage();
            var app = page1.Application;
            var docs = app.Documents;
            var doc = page1.Document;
            var colors = doc.Colors;

            check_col( docs.Cast<IVisio.Document>().ToList() , docs.AsEnumerable().ToList() );
            check_col( colors.Cast<IVisio.Color>().ToList(), colors.AsEnumerable().ToList(), (a,b) => a.Red==b.Red && a.Green==b.Green && a.Blue ==b.Blue);
        
    

        }

        private void check_col<T>(IList<T> expected, IList<T> actual, System.Func<T,T,bool> f) where T : class
        {
            if (expected.Count != actual.Count)
            {
                Assert.Fail("sizes do not match");
            }

            for (int i = 0; i < expected.Count; i++)
            {
                T ex = expected[i];
                T ac = actual[i];
                if (!f(ex,ac))
                {
                    Assert.Fail("objects don't match");
                }
            }
        }
        private void check_col<T> (IList<T> expected, IList<T> actual) where T: class
        {
            check_col(expected,actual, (a,b) => a!=b);
        }

    }
}