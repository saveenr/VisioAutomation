using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenTreeOps_Test
{
    [TestClass]
    public class PreOrderTests
    {
        [TestMethod]
        public void Walk_4()
        {
            var n0 = new XNode("A");

            var preorder_results = GenTreeOps.Algorithms.PreOrder(n0, n => n.Children).ToList();
            var preorder_string = string.Join("", preorder_results.Select(n => n.Name));

            Assert.AreEqual("A", preorder_string);
        }

        [TestMethod]
        public void Walk_5()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            n0.Children.Add(n1);

            var preorder_results = GenTreeOps.Algorithms.PreOrder(n0, n => n.Children).ToList();
            var preorder_string = string.Join("", preorder_results.Select(n => n.Name));

            Assert.AreEqual("AB", preorder_string);
        }

        [TestMethod]
        public void Walk_6()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            var n2 = new XNode("C");
            var n3 = new XNode("D");
            n0.Children.Add(n1);
            n0.Children.Add(n2);
            n2.Children.Add(n3);

            var preorder_results = GenTreeOps.Algorithms.PreOrder(n0, n => n.Children).ToList();
            var preorder_string = string.Join("", preorder_results.Select(n => n.Name));

            Assert.AreEqual("ABCD", preorder_results.Count);
        }

    }
}