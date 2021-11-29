using System.Collections.Generic;
using System.Linq;
using GenTreeOps;

namespace VisioAutomation.Models.Layouts.InternalTree
{
    internal class Node<T>
    {
        private List<Node<T>> _child_list;

        private bool _is_collapsed;
        private static int _nodeSeqNum = 0;

        internal double Modifier; // temporary modifier
        internal double PrelimX; // preliminary x coord    
        internal Node<T> LeftNeighbor;
        internal Node<T> RightNeighbor;

        public int Id { get; set; }

        public VisioAutomation.Core.Size Size { get; set; }

        public VisioAutomation.Core.Rectangle Rect => new VisioAutomation.Core.Rectangle(this.Position, this.Size);

        internal void Init(int id, Node<T> parent, VisioAutomation.Core.Size size, T data)
        {
            this.Id = id;
            this.Size = size;
            this.Data = data;
            this.Parent = parent;

            this._child_list = new List<Node<T>>();
            this.LeftNeighbor = null;
            this.RightNeighbor = null;
            this.Position = new VisioAutomation.Core.Point(0, 0);
            this._is_collapsed = false;
        }

        internal Node(int id, Node<T> parent, VisioAutomation.Core.Size size)
        {
            this.Init(id, parent, size, default(T));
        }

        public Node(VisioAutomation.Core.Size size, T data)
        {
            this.Init(Node<T>._nodeSeqNum++, null, size, data);
        }

        public int ChildCount
        {
            get
            {
                if (this._is_collapsed)
                {
                    return 0;
                }
                if (this._child_list == null)
                {
                    return 0;
                }
                return this._child_list.Count;
            }
        }

        public Node<T> LeftSibling
        {
            get
            {
                if (this.LeftNeighbor != null && this.LeftNeighbor.Parent == this.Parent)
                {
                    return this.LeftNeighbor;
                }
                return null;
            }
        }

        public Node<T> RightSibling
        {
            get
            {
                if (this.RightNeighbor != null && this.RightNeighbor.Parent == this.Parent)
                {
                    return this.RightNeighbor;
                }
                return null;
            }
        }

        public Node<T> FirstChild => this.GetChildAt(0);

        public Node<T> LastChild => this.GetChildAt(this.ChildCount - 1);

        private void add_child(Node<T> nn)
        {
            nn.Parent = this;
            this._child_list.Add(nn);
        }

        public Node<T> AddChild(Node<T> child)
        {
            this.add_child(child);
            return child;
        }

        public Node<T> AddNewChild(VisioAutomation.Core.Size size)
        {
            var new_child = new Node<T>(Node<T>._nodeSeqNum++, null, size);
            this.add_child(new_child);
            return new_child;
        }

        public int Level
        {
            get
            {
                if (this.Parent.Id == -1)
                {
                    return 0;
                }
                return this.Parent.Level + 1;
            }
        }

        public Node<T> Parent { get; set; }

        public T Data { get; set; }

        public VisioAutomation.Core.Point Position { get; set; }

        public bool GetIsAncestorCollapsed()
        {
            if (this.Parent._is_collapsed)
            {
                return true;
            }
            if (this.Parent.Id == -1)
            {
                return false;
            }
            return this.Parent.GetIsAncestorCollapsed();
        }

        public Node<T> GetChildAt(int index)
        {
            return this._child_list[index];
        }

        public double GetChildrenCenter(TreeLayout<T> tree_layout_engine)
        {
            var node0 = this.FirstChild;
            var node1 = this.LastChild;
            return node0.PrelimX + ((node1.PrelimX - node0.PrelimX) + tree_layout_engine.GetNodeSize(node1))/2;
        }

        public IEnumerable<Node<T>> EnumChildren()
        {
            foreach (var c in this._child_list)
            {
                yield return c;
            }
        }

        public IEnumerable<Node<T>> EnumRecursive()
        {
            var iter = GenTreeOps.Algorithms.Walk<Node<T>>(this, n => n.EnumChildren());
            var iter2 = iter.Where(i => i.Type == WalkEventType.EventEnter).Select(i => i.Node);
            return iter2;
        }
    }
}