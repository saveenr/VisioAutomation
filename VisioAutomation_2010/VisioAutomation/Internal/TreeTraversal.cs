using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Internal
{
    internal static class TreeTraversal
    {
        // Delegates
        public delegate IEnumerable<T> EnumerateChildren<T>(T item);

        private struct WalkState<T>
        {
            // this is an internal struct used when traversing the DOM
            // it preserves the state in the non-recursive, stack-based traversal of the DOM

            internal readonly T Node;
            internal bool Entered;

            public WalkState(T node)
            {
                this.Node = node;
                this.Entered = false;
            }
        }

        /// <summary>
        // Walks a Node in a depth-first/pre-order manner without recursion.
        // It returns a series of "events" that indicate one of three things:
        // - whether it has enters into a node
        // - whether it has exited from a node (i.e. it is finished with that container and its children)
        // - caller can control which children get entered via the enum_children method
        /// </summary>
        public static IEnumerable<WalkEvent<T>> Walk<T>(T node, TreeTraversal.EnumerateChildren<T> enum_children)
        {
            var stack = new Stack<WalkState<T>>();

            // put the first item on the stack 
            stack.Push(new WalkState<T>(node));

            // As long as something is on the stack, we are not done
            while (stack.Count > 0)
            {
                var cur_item = stack.Pop();

                if (cur_item.Entered == false)
                {
                    var walkevent = WalkEvent<T>.CreateEnterEvent(cur_item.Node);
                    yield return walkevent;

                    cur_item.Entered = true;
                    stack.Push(cur_item);

                    foreach (var child in TreeTraversal.efficient_reverse(enum_children(cur_item.Node)))
                    {
                        stack.Push(new WalkState<T>(child));
                    }
                }
                else
                {
                    var walkevent = WalkEvent<T>.CreateExitEvent(cur_item.Node);
                    yield return walkevent;
                }
            }
        }

        public static IEnumerable<T> PreOrder<T>(T root, EnumerateChildren<T> enum_children)
        {
            foreach (var ev in Walk(root,enum_children))
            {
                if (ev.HasEnteredNode)
                {
                    yield return ev.Node;
                }
            }
        }

        internal static IEnumerable<T> efficient_reverse<T>(IEnumerable<T> items)
        {
            if (items is IList<T>)
            {
                var item_col = (IList<T>) items;
                for (int i = item_col.Count - 1; i >= 0; i--)
                {
                    yield return item_col[i];
                }
            }
            else
            {
                foreach (var i in items.Reverse())
                {
                    yield return i;
                }
            }
        }
    }
}