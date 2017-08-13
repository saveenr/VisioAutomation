using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenTreeOps_Test
{
    [TestClass]
    public class PostOrderTests
    {


        [TestMethod]
        public void PreOrder_1()
        {
            // tree comes from: https://stackoverflow.com/questions/9456937/when-to-use-preorder-postorder-and-inorder-binary-search-tree-traversal-strate

            var n7 = new XNode("7");

            var n1 = new XNode("1");
            n7.Children.Add(n1);

            var n0 = new XNode("0");
            n1.Children.Add(n0);

            var n3 = new XNode("3");
            n1.Children.Add(n3);

            var n2 = new XNode("2");
            n3.Children.Add(n2);

            var n5 = new XNode("5");
            n3.Children.Add(n5);

            var n4 = new XNode("4");
            n5.Children.Add(n4);

            var n6 = new XNode("6");
            n5.Children.Add(n6);

            var n9 = new XNode("9");
            n7.Children.Add(n9);

            var n8 = new XNode("8");
            n9.Children.Add(n8);

            var n10 = new XNode("10");
            n9.Children.Add(n10);

            var preorder_results = GenTreeOps.Algorithms.PreOrder(n7, n => n.Children).ToList();
            var preorder_string = string.Join(", ", preorder_results.Select(n => n.Name));

            Assert.AreEqual("7, 1, 0, 3, 2, 5, 4, 6, 9, 8, 10", preorder_string);

            var postorder_results = GenTreeOps.Algorithms.PostOrder(n7, n => n.Children).ToList();
            var postorder_string = string.Join(", ", postorder_results.Select(n => n.Name));

            Assert.AreEqual("0, 2, 4, 6, 5, 3, 1, 8, 10, 9, 7", postorder_string);

        }
    }
}