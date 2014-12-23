using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Models.InternalTree
{
    internal class Node<T>
    {
        private List<Node<T>> child_list;

        private bool is_collapsed;
        private static int node_seq_num = 0;

        internal double modifier; // temporary modifier
        internal double prelim_x; // preliminary x coord    
        internal Node<T> left_neighbor;
        internal Node<T> right_neighbor;

        public int Id { get; set; }

        public VA.Drawing.Size Size { get; set; }

        public VA.Drawing.Rectangle Rect
        {
            get { return new VA.Drawing.Rectangle(this.Position, this.Size); }
        }

        internal void init(int id, Node<T> parent, VA.Drawing.Size size, T data)
        {
            this.Id = id;
            this.Size = size;
            this.Data = data;
            this.Parent = parent;

            child_list = new List<Node<T>>();
            left_neighbor = null;
            right_neighbor = null;
            Position = new VA.Drawing.Point(0, 0);
            is_collapsed = false;
        }

        internal Node(int id, Node<T> parent, VA.Drawing.Size size)
        {
            init(id, parent, size, default(T));
        }

        public Node(VA.Drawing.Size size, T data)
        {
            init(node_seq_num++, null, size, data);
        }

        public int ChildCount
        {
            get
            {
                if (is_collapsed)
                {
                    return 0;
                }
                if (child_list == null)
                {
                    return 0;
                }
                return child_list.Count;
            }
        }

        public Node<T> LeftSibling
        {
            get
            {
                if (left_neighbor != null && left_neighbor.Parent == Parent)
                {
                    return left_neighbor;
                }
                return null;
            }
        }

        public Node<T> RightSibling
        {
            get
            {
                if (right_neighbor != null && right_neighbor.Parent == Parent)
                {
                    return right_neighbor;
                }
                return null;
            }
        }

        public Node<T> FirstChild
        {
            get { return GetChildAt(0); }
        }

        public Node<T> LastChild
        {
            get { return GetChildAt(ChildCount - 1); }
        }

        private void add_child(Node<T> nn)
        {
            nn.Parent = this;
            this.child_list.Add(nn);
        }

        public Node<T> AddChild(Node<T> child)
        {
            this.add_child(child);
            return child;
        }

        public Node<T> AddNewChild(VA.Drawing.Size size)
        {
            var new_child = new Node<T>(node_seq_num++, null, size);
            this.add_child(new_child);
            return new_child;
        }

        public int Level
        {
            get
            {
                if (Parent.Id == -1)
                {
                    return 0;
                }
                return Parent.Level + 1;
            }
        }

        public Node<T> Parent { get; set; }

        public T Data { get; set; }

        public VA.Drawing.Point Position { get; set; }

        public bool GetIsAncestorCollapsed()
        {
            if (Parent.is_collapsed)
            {
                return true;
            }
            if (Parent.Id == -1)
            {
                return false;
            }
            return Parent.GetIsAncestorCollapsed();
        }

        public Node<T> GetChildAt(int index)
        {
            return child_list[index];
        }

        public double GetChildrenCenter(TreeLayout<T> treeLayoutEngine)
        {
            var node0 = FirstChild;
            var node1 = LastChild;
            return node0.prelim_x + ((node1.prelim_x - node0.prelim_x) + treeLayoutEngine.GetNodeSize(node1))/2;
        }

        public IEnumerable<Node<T>> EnumChildren()
        {
            foreach (var c in child_list)
            {
                yield return c;
            }
        }

        public IEnumerable<Node<T>> EnumRecursive()
        {
            var iter = VA.Internal.TreeOps.Walk<Node<T>>(this, n => n.EnumChildren());
            var iter2 = iter.Where(i => i.HasEnteredNode).Select(i => i.Node);
            return iter2;
        }
    }
}