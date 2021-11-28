using System.Collections.Generic;
using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.DocumentAnalysis
{
    [MUT.TestClass]
    public class Path_Test
    {
        [MUT.TestMethod]
        public void Path_TestTransitiveClosure0()
        {
            // v0->v0
            // doesn't yield any edges (nodes are implictly connected to themselves)
            var input = new List<VisioAutomation.Analyzers.DirectedEdge<string, object>>
                {
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v0", "v0", null)
                };
            var output = VisioAutomation.Analyzers.ConnectionAnalyzer.GetClosureFromEdges(input).ToList();
            MUT.Assert.AreEqual(0,output.Count);
        }

        [MUT.TestMethod]
        public void Path_TestTransitiveClosure1()
        {
            // v0->v0
            // v1->v1
            // doesn't yield any edges (nodes are implictly connected to themselves)
            var input = new List<VisioAutomation.Analyzers.DirectedEdge<string, object>>();
            input.Add(new VisioAutomation.Analyzers.DirectedEdge<string, object>("v0", "v0", null));
            input.Add(new VisioAutomation.Analyzers.DirectedEdge<string, object>("v1", "v1", null));
            var output = VisioAutomation.Analyzers.ConnectionAnalyzer.GetClosureFromEdges(input).ToList();
            MUT.Assert.AreEqual(0, output.Count);
        }

        [MUT.TestMethod]
        public void Path_TestTransitiveClosure2()
        {
            // v0->v1
            // doesn't yield any edges (nodes are implictly connected to themselves)
            var input = new List<VisioAutomation.Analyzers.DirectedEdge<string, object>>
                {
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v0", "v1", null)
                };
            var output = VisioAutomation.Analyzers.ConnectionAnalyzer.GetClosureFromEdges(input).ToList();
            MUT.Assert.AreEqual(1, output.Count);
            MUT.Assert.AreEqual("v0",output[0].From);
            MUT.Assert.AreEqual("v1", output[0].To);
        }


        [MUT.TestMethod]
        public void Path_TestTransitiveClosure3()
        {
            var input = new List<VisioAutomation.Analyzers.DirectedEdge<string, object>>
                {
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v0", "v1", null),
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v1", "v2", null)
                };
            var output = VisioAutomation.Analyzers.ConnectionAnalyzer.GetClosureFromEdges(input).ToList();
            MUT.Assert.AreEqual(3, output.Count);
            MUT.Assert.AreEqual("v0", output[0].From);
            MUT.Assert.AreEqual("v1", output[0].To);

            MUT.Assert.AreEqual("v0", output[1].From);
            MUT.Assert.AreEqual("v2", output[1].To);

            MUT.Assert.AreEqual("v1", output[2].From);
            MUT.Assert.AreEqual("v2", output[2].To);
            
        }

        [MUT.TestMethod]
        public void Path_TestTransitiveClosure4()
        {
            var input = new List<VisioAutomation.Analyzers.DirectedEdge<string, object>>
                {
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v0", "v1", null),
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v1", "v2", null),
                    new VisioAutomation.Analyzers.DirectedEdge<string, object>("v2", "v0", null)
                };
            var output = VisioAutomation.Analyzers.ConnectionAnalyzer.GetClosureFromEdges(input).ToList();
            MUT.Assert.AreEqual(6, output.Count);
            MUT.Assert.AreEqual("v0", output[0].From);
            MUT.Assert.AreEqual("v1", output[0].To);

            MUT.Assert.AreEqual("v0", output[1].From);
            MUT.Assert.AreEqual("v2", output[1].To);

            MUT.Assert.AreEqual("v1", output[2].From);
            MUT.Assert.AreEqual("v0", output[2].To);

            MUT.Assert.AreEqual("v1", output[3].From);
            MUT.Assert.AreEqual("v2", output[3].To);

            MUT.Assert.AreEqual("v2", output[4].From);
            MUT.Assert.AreEqual("v0", output[4].To);

            MUT.Assert.AreEqual("v2", output[5].From);
            MUT.Assert.AreEqual("v1", output[5].To);

        }
    }
}