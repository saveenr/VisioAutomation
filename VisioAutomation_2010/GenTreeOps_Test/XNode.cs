using System.Collections.Generic;

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
    }
}