using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class PathAnalysis_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Case1()
        {
            var page1 = GetNewPage();
            var stencil = page1.Application.Documents.OpenStencil("basic_u.vss");
            var dcm = stencil.Masters["Dynamic Connector"];

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(2, 0, 3, 1);
            var c1 = page1.Drop(dcm, new VA.Drawing.Point(-2, -2));

            VA.Connections.ConnectorHelper.ConnectShapes(c1, s1, s2);

            var edges_bd = VA.Connections.PathAnalysis.GetEdges(page1,
                                                             VA.Connections.ConnectorArrowEdgeHandling.
                                                                 TreatNoArrowEdgesAsBidirectional);
            var edges_d = VA.Connections.PathAnalysis.GetEdges(page1,
                                                             VA.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);

            var tcbd = VA.Connections.PathAnalysis.GetTransitiveClosure(page1,
                                                             VA.Connections.ConnectorArrowEdgeHandling.
                                                                 TreatNoArrowEdgesAsBidirectional);

            var tcd = VA.Connections.PathAnalysis.GetTransitiveClosure(page1,
                                                 VA.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);


            page1.Delete(0);
        }

    }
}