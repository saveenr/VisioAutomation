using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using VisioAutomation.Extensions;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class ConnectorTests : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void Connect1()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(1, 1, 2, 2);
            var s2 = page1.DrawRectangle(5, 5, 6, 6);
            var stencil = page1.Application.Documents.OpenStencil("connec_u.vss");
            var master = stencil.Masters["Dynamic Connector"];
            var connector = page1.Drop(master, 0, 0);
            ConnectorHelper.ConnectShapes(s1, s2, connector);

            page1.Delete(0);
        }
    }
}