using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Internal
{
    internal static class TreeUtil
    {
        public static IList<TDest> TransformTree<TSrc, TDest>(
            TSrc rootnode,
            System.Func<TSrc, IEnumerable<TSrc>> enumchildren,
            System.Func<TSrc, TDest> createdstnode,
            System.Action<TDest, TDest> addchild)
        {
            var stack = new Stack<TDest>();
            var tnodes = new List<TDest>();

            var walk_items = VA.Internal.TreeTraversal.Walk<TSrc>(rootnode, input_node => enumchildren(input_node));
            foreach (var walk_item in walk_items)
            {
                if (walk_item.HasEnteredNode)
                {
                    var new_dst_node = createdstnode(walk_item.Node);

                    if (stack.Count > 0)
                    {
                        var parent = stack.Peek();
                        addchild(parent, new_dst_node);
                    }
                    stack.Push(new_dst_node);
                    tnodes.Add(new_dst_node);
                }
                else if (walk_item.HasExitedNode)
                {
                    stack.Pop();
                }
            }

            return tnodes;
        }    
    }
}