using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio= Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VTest.Core.Extensions
{
    [MUT.TestClass]
    public class PageDropTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Page_Drop_ManyU()
        {
            var page1 = this.GetNewPage();            
            var stencil = "basic_u.vss";

            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var app = page1.Application;
            var documents = app.Documents;
            var stencil_doc = documents.OpenEx(stencil, flags);

            var masters1 = stencil_doc.Masters;
            var masters = new [] {masters1["Rounded Rectangle"], masters1["Ellipse"]};
            var points = new [] {new VA.Core.Point(1, 2), new VA.Core.Point(3, 4)};
            MUT.Assert.AreEqual(0, page1.Shapes.Count);
            var shapeids = page1.DropManyU(masters, points);
            MUT.Assert.AreEqual(2, page1.Shapes.Count);
            MUT.Assert.AreEqual(2, shapeids.Length );

            var s0 = page1.Shapes[shapeids[0]];
            var s1 = page1.Shapes[shapeids[1]];

            MUT.Assert.AreEqual( masters[0].NameU, s0.Master.NameU );
            MUT.Assert.AreEqual(masters[1].NameU, s1.Master.NameU);
            
            page1.Delete(0);
        }
    }
}
