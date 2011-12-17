using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Internal
{
    internal static class TreeUtil
    {
        public static IList<TDest> CopyTree<TSrc, TDest>(
            TSrc src_root_node,
            System.Func<TSrc, IEnumerable<TSrc>> enum_src_children,
            System.Func<TSrc, TDest> create_dest_node,
            System.Action<TDest, TDest> add_dest_child)
        {
            var stack = new Stack<TDest>();
            var dest_nodes = new List<TDest>();

            var walk_items = VA.Internal.TreeTraversal.Walk<TSrc>(src_root_node, input_node => enum_src_children(input_node));
            foreach (var walk_item in walk_items)
            {
                if (walk_item.HasEnteredNode)
                {
                    var new_dst_node = create_dest_node(walk_item.Node);

                    if (stack.Count > 0)
                    {
                        // if there is node on the stack, then that node is the current node's parent
                        var parent = stack.Peek();
                        add_dest_child(parent, new_dst_node);
                    }
                    else
                    {
                        // if there is nothing on the stack this is node without a parent (a root node)
                    }

                    stack.Push(new_dst_node);
                    dest_nodes.Add(new_dst_node);
                }
                else if (walk_item.HasExitedNode)
                {
                    stack.Pop();
                }
            }

            return dest_nodes;
        }    
    }
}