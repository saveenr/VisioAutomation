using System.Collections.Generic;
using System.Linq;

namespace GenTreeOps_Test
{
    public class XNode
    {
        public readonly string Name;
        public readonly List<XNode> Children;

        public XNode(string name)
        {
            this.Name = name;
            this.Children = new List<XNode>();
        }

        public string GetPreorderString()
        {
            var preorder_results = GenTreeOps.Algorithms.PreOrder(this, n => n.Children).ToList();
            var preorder_string = string.Join("", preorder_results.Select(n => n.Name));
            return preorder_string;
        }
    }
}