using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class PerfTests : VisioAutomationTest
    {
        [TestMethod]
        public void Perf1()
        {
            var page1 = GetNewPage();
            var shape0 = page1.DrawRectangle(1, 1, 3, 3);
            int n = 10;
            var t1 = new System.Diagnostics.Stopwatch();
            t1.Start();
            for (int i=0;i<n;i++)
            {
                VisioAutomationTest.GetSize(shape0);
            }
            t1.Stop();
            page1.Delete(0);
        }
    }
}