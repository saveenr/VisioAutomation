using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenTreeOps_Test
{
    [TestClass]
    public class CopyTests
    {
        [TestMethod]
        public void Walk_4()
        {
            var n0 = new XNode("A");

            var output = GenTreeOps.Algorithms.CopyTree(
                n0, // the source root 
                src_n => src_n.Children, // how to enum src children
                src_n => new XNode(src_n.Name), // how to create a dest node
                (dest_p, dest_c) => dest_p.Children.Add(dest_c) // how ot add a child
            );

            string src = n0.GetPreorderString();
            string dest = output[0].GetPreorderString();

            Assert.AreEqual("A", src);
            Assert.AreEqual("A", dest);
        }

        [TestMethod]
        public void Walk_5()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            n0.Children.Add(n1);

            var output = GenTreeOps.Algorithms.CopyTree(
                n0, // the source root 
                src_n => src_n.Children, // how to enum src children
                src_n => new XNode(src_n.Name), // how to create a dest node
                (dest_p, dest_c) => dest_p.Children.Add(dest_c) // how ot add a child
            );

            string src = n0.GetPreorderString();
            string dest = output[0].GetPreorderString();

            Assert.AreEqual("AB", src);
            Assert.AreEqual("AB", dest);
        }

        [TestMethod]
        public void Copy_3()
        {
            var n0 = new XNode("A");
            var n1 = new XNode("B");
            var n2 = new XNode("C");
            var n3 = new XNode("D");
            n0.Children.Add(n1);
            n0.Children.Add(n2);
            n2.Children.Add(n3);

            var output = GenTreeOps.Algorithms.CopyTree(
                n0, // the source root 
                src_n => src_n.Children, // how to enum src children
                src_n => new XNode(src_n.Name), // how to create a dest node
                (dest_p,dest_c) => dest_p.Children.Add(dest_c) // how ot add a child
            );

            string src = n0.GetPreorderString();
            string dest = output[0].GetPreorderString();

            Assert.AreEqual("ABCD", src);
            Assert.AreEqual("ABCD", dest);
        }

    }
}