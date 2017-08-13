using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenTreeOps_Test
{
    [TestClass]
    public class WalkTests
    {
        [TestMethod]
        public void Walk_1()
        {
            var n0 = new XNode("A");
            var events = GenTreeOps.Algorithms.Walk(n0, n => n.Children).ToList();

            AssertX.AssertEnter(n0, events[0]);
            AssertX.AssertExit(n0, events[1]);
        }

        [TestMethod]
        public void Walk_2()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            n0.Children.Add(n1);

            var events = GenTreeOps.Algorithms.Walk(n0, n => n.Children).ToList();

            Assert.AreEqual(4, events.Count);

            AssertX.AssertEnter(n0, events[0]);
            AssertX.AssertEnter(n1, events[1]);
            AssertX.AssertExit(n1, events[2]);
            AssertX.AssertExit(n0, events[3]);
        }

        [TestMethod]
        public void Walk_3()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            var n2 = new XNode("C");
            var n3 = new XNode("D");
            n0.Children.Add(n1);
            n0.Children.Add(n2);
            n2.Children.Add(n3);

            var events = GenTreeOps.Algorithms.Walk(n0, n => n.Children).ToList();

            Assert.AreEqual(4 * 2, events.Count);


            // enter A
            AssertX.AssertEnter(n0, events[0]);
            AssertX.AssertEnter(n1, events[1]);
            AssertX.AssertExit(n1, events[2]);
            AssertX.AssertEnter(n2, events[3]);
            AssertX.AssertEnter(n3, events[4]);
            AssertX.AssertExit(n3, events[5]);
            AssertX.AssertExit(n2, events[6]);
            AssertX.AssertExit(n0, events[7]);
        }
    }
}
