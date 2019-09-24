using System.Collections.Generic;
using System.Linq;


/*
 * 
 * This is a C# translation of the JavaScript source code of "Graphic JavaScript LayoutTree with Layout" by Emilio Cortegoso Lobato 
 * http://www.codeproject.com/KB/scripting/graphic_javascript_tree.aspx
 * 
 * 
 * That code is in turn is based on: "Positioning Nodes For General Trees" by John Q. Walker II
 * This algorithm has been published in several locations:
 * 1. http://www.cs.unc.edu/techreports/89-034.pdf 
 * 2. "Software - Practice and Experience", July 1990, Copyright 1990 by John Wiley and Sons, Ltd.
 * 
 * Other literature:
 * - Buchheim 
 *   http://dirk.jivas.de/papers/buchheim02improving.pdf
 *   https://github.com/d3/d3-hierarchy/blob/master/src/tree.js
 * - Tidy Drawings of Trees (Charles Wetherell & Alfred Shannon)
 *   http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.150.4061&rep=rep1&type=pdf
 * - Other 
 *   https://llimllib.github.io/pymag-trees/
 *   https://github.com/fforw/rt03-mindmap/blob/master/src/script/tree-layout.js
 *
 * * KEY UPDATES
 * ------------
 * - C# Implementation
 * - Separated formatting from layout and removed formatting information
 * - Strongly typed
 * - Works with the origin in the lower left
 * - uses VA structs such as Rectangle, Point, Size
 * - Names of methods changed to match guidelines for .NET Libraries
 * - Added back in some comments from the original source code by John Q. Walker II
 * 
 * */

namespace VisioAutomation.Models.Layouts.InternalTree
{
    internal class TreeLayout<T>
    {
        private Dictionary<int, double> _max_level_height;
        private Dictionary<int, double> _max_level_width;
        private Dictionary<int, Node<T>> _previous_level_node;
        private VisioAutomation.Geometry.Point _root_offset;
        private readonly Node<T> _root;

        public TreeLayoutOptions Options { get; set; }

        public Node<T> Root => this._root;

        public TreeLayout()
        {
            this.Options = new TreeLayoutOptions();
            this._root = new Node<T>(-1, null, this.Options.DefaultNodeSize);
        }

        public IEnumerable<Node<T>> Nodes => this.Root.EnumRecursive().Skip(1); // return all the nodes (except the special root)

        private void set_level_height(Node<T> node, int level)
        {
            var value = this._max_level_height.GetValueOrDefaultEx( level, 0);
            this._max_level_height[level] = System.Math.Max(value, node.Size.Height);
        }

        private void set_level_width(Node<T> node, int level)
        {
            var value = this._max_level_width.GetValueOrDefaultEx(level, 0);
            this._max_level_width[level] = System.Math.Max(value, node.Size.Width);
        }

        private void set_neighbors(Node<T> node, int level)
        {
            node.LeftNeighbor = this._previous_level_node.GetValueOrDefaultEx(level, null);

            if (node.LeftNeighbor != null)
            {
                node.LeftNeighbor = this._previous_level_node[level];
                node.LeftNeighbor.RightNeighbor = node;
            }

            this._previous_level_node[level] = node;
        }

        public double GetNodeSize(Node<T> node)
        {
            return this.Options.Direction switch
            {
                LayoutDirection.Up => node.Size.Width,
                LayoutDirection.Left => node.Size.Height,
                LayoutDirection.Right => node.Size.Height,
                LayoutDirection.Down => node.Size.Width,
                _ => throw new System.ArgumentOutOfRangeException(),
            };
        }

        private void _apportion(Node<T> node, int level)
        {
            /*------------------------------------------------------
             * Clean up the positioning of small sibling subtrees.
             * Subtrees of a node are formed independently and
             * placed as close together as possible. By requiring
             * that the subtrees be rigid at the time they are put
             * together, we avoid the undesirable effects that can
             * accrue from positioning nodes rather than subtrees.
             *----------------------------------------------------*/

            var first_child = node.FirstChild;
            var first_child_left_neighbor = first_child.LeftNeighbor;
            int j = 1;
            for (int k = this.Options.MaximumDepth - level;
                 first_child != null && first_child_left_neighbor != null && j <= k;)
            {
                double modifier_sum_right = 0;
                double modifier_sum_left = 0;
                var right_ancestor = first_child;
                var left_ancestor = first_child_left_neighbor;
                for (int l = 0; l < j; l++)
                {
                    right_ancestor = right_ancestor.Parent;
                    left_ancestor = left_ancestor.Parent;
                    modifier_sum_right += right_ancestor.Modifier;
                    modifier_sum_left += left_ancestor.Modifier;
                }

                double total_gap = (first_child_left_neighbor.PrelimX + modifier_sum_left +
                                   this.GetNodeSize(first_child_left_neighbor) + this.Options.SubtreeSeparation) -
                                  (first_child.PrelimX + modifier_sum_right);
                if (total_gap > 0)
                {
                    var subtree_aux = node;
                    int num_subtrees = 0;
                    for (; subtree_aux != null && subtree_aux != left_ancestor; subtree_aux = subtree_aux.LeftSibling)
                    {
                        num_subtrees++;
                    }

                    if (subtree_aux != null)
                    {
                        var subtree_move_aux = node;
                        double single_gap = total_gap/num_subtrees;
                        for (; subtree_move_aux != left_ancestor; subtree_move_aux = subtree_move_aux.LeftSibling)
                        {
                            subtree_move_aux.PrelimX += total_gap;
                            subtree_move_aux.Modifier += total_gap;
                            total_gap -= single_gap;
                        }
                    }
                }
                j++;

                if (first_child.ChildCount == 0)
                {
                    first_child = TreeLayout<T>.get_leftmost(node, 0, j);
                }
                else
                {
                    first_child = first_child.FirstChild;
                }

                if (first_child != null)
                {
                    first_child_left_neighbor = first_child.LeftNeighbor;
                }
            }
        }

        public VisioAutomation.Geometry.Rectangle GetBoundingBoxOfTree()
        {
            if (this.Root.ChildCount < 1)
            {
                throw new System.InvalidOperationException("There are no Nodes in the tree");
            }
            var nodes = this.Nodes.ToList();

            var bb = Geometry.BoundingBoxBuilder.FromRectangles(nodes.Select(n => n.Rect));
            if (!bb.HasValue)
            {
                throw new System.InvalidOperationException("Internal Error: Could not compute bounding box");
            }
            else
            {
                return bb.Value;
            }
        }

        private static Node<T> get_leftmost(Node<T> node, int level, int maxlevel)
        {
            if (level >= maxlevel)
            {
                return node;
            }
            if (node.ChildCount == 0)
            {
                return null;
            }

            foreach (var child in node.EnumChildren())
            {
                var leftmost_descendant = TreeLayout<T>.get_leftmost(child, level + 1, maxlevel);
                if (leftmost_descendant != null)
                {
                    return leftmost_descendant;
                }
            }

            return null;
        }

        //Layout algorithm
        private void first_walk(Node<T> node, int level)
        {
            node.Position = new VisioAutomation.Geometry.Point(0, 0);
            node.PrelimX = 0;
            node.Modifier = 0;
            node.LeftNeighbor = null;
            node.RightNeighbor = null;
            this.set_level_height(node, level);
            this.set_level_width(node, level);
            this.set_neighbors(node, level);
            if (node.ChildCount == 0 || level == this.Options.MaximumDepth)
            {
                var left_sibling = node.LeftSibling;
                if (left_sibling != null)
                {
                    /*--------------------------------------------
                     * Determine the preliminary x-coordinate
                     *   based on:
                     * - preliminary x-coordinate of left sibling,
                     * - the separation between sibling nodes, and
                     * - mean width of left sibling & current node.
                     *--------------------------------------------*/

                    node.PrelimX = left_sibling.PrelimX + this.GetNodeSize(left_sibling) +
                                    this.Options.SiblingSeparation;
                }
                else
                {
                    /*  no sibling on the left to worry about  */

                    node.PrelimX = 0;
                }
            }
            else
            {
                /* Position the leftmost of the children          */

                foreach (var child  in node.EnumChildren())
                {
                    this.first_walk(child, level + 1);
                }

                /* Calculate the preliminary value between   */
                /* the children at the far left and right    */

                double mid_point = node.GetChildrenCenter(this) - this.GetNodeSize(node)/2.0;

                if (node.LeftSibling != null)
                {
                    node.PrelimX = node.LeftSibling.PrelimX +
                                    this.GetNodeSize(node.LeftSibling) +
                                    this.Options.SiblingSeparation;
                    node.Modifier = node.PrelimX - mid_point;
                    this._apportion(node, level);
                }
                else
                {
                    node.PrelimX = mid_point;
                }
            }
        }

        private void second_walk(Node<T> node, int level, VisioAutomation.Geometry.Point p)
        {
            /*------------------------------------------------------
                * During a second pre-order walk, each node is given a
                * final x-coordinate by summing its preliminary
                * x-coordinate and the modifiers of all the node's
                * ancestors.  The y-coordinate depends on the height of
                * the tree.  (The roles of x and y are reversed for
                * RootOrientations of EAST or WEST.)
                * Returns: TRUE if no errors, otherwise returns FALSE.
                *----------------------------------------- ----------*/

            if (level > this.Options.MaximumDepth) return;

            var temp_point = this._root_offset.Add(node.PrelimX, 0) + p;
            double maxsize_tmp = 0;
            double nodesize_tmp = 0;
            bool flag = false;

            switch (this.Options.Direction)
            {
                case LayoutDirection.Up:
                case LayoutDirection.Down:
                    {
                        maxsize_tmp = this._max_level_height[level];
                        nodesize_tmp = node.Size.Height;
                        break;
                    }
                case LayoutDirection.Left:
                case LayoutDirection.Right:
                    {
                        maxsize_tmp = this._max_level_width[level];
                        flag = true;
                        nodesize_tmp = node.Size.Width;
                        break;
                    }
            }
            switch (this.Options.Alignment)
            {
                case AlignmentVertical.Top:
                    node.Position = temp_point;
                    break;

                case AlignmentVertical.Center:
                    node.Position = temp_point.Add(0, (maxsize_tmp - nodesize_tmp)/2.0);
                    break;

                case AlignmentVertical.Bottom:
                    node.Position = temp_point.Add(0, maxsize_tmp - nodesize_tmp);
                    break;
            }

            if (flag)
            {
                // QUESTION: Why is this step performed?
                node.Position = new VisioAutomation.Geometry.Point(node.Position.Y, node.Position.X);
            }

            switch (this.Options.Direction)
            {
                case LayoutDirection.Down:
                    {
                        node.Position = new VisioAutomation.Geometry.Point(node.Position.X, -node.Position.Y - nodesize_tmp);
                        break;
                    }
                case LayoutDirection.Left:
                    {
                        node.Position = new VisioAutomation.Geometry.Point(-node.Position.X - nodesize_tmp, node.Position.Y);
                        break;
                    }
            }

            if (node.ChildCount != 0)
            {
                /* Apply the flModifier value for this    */
                /* node to all its offspring.             */

                var np = p.Add(node.Modifier, maxsize_tmp + this.Options.LevelSeparation);
                this.second_walk(node.FirstChild, level + 1, np);
            }

            if (node.RightSibling != null)
            {
                this.second_walk(node.RightSibling, level, p);
            }
        }

        public void PerformLayout()
        {
            /*------------------------------------------------------
             * Determine the coordinates for each node in a tree.
             * Input: Pointer to the apex node of the tree
             * Assumption: The x & y coordinates of the apex node
             * are already correct, since the tree underneath it
             * will be positioned with respect to those coordinates.
             * Returns: TRUE if no errors, otherwise returns FALSE.
             *----------------------------------------------------*/

            this._max_level_height = new Dictionary<int, double>();
            this._max_level_width = new Dictionary<int, double>();
            this._previous_level_node = new Dictionary<int, Node<T>>();

            this.first_walk(this._root, 0);

            //adjust the root_offset
            // NOTE: in the original code this was a case statement on Options.Direction that did the same thing for each direction 
            this._root_offset = this.Options.TopAdjustment + this._root.Position;

            this.second_walk(this._root, 0, new VisioAutomation.Geometry.Point(0, 0));

            this._max_level_height = null;
            this._max_level_width = null;
            this._previous_level_node = null;

            this.correct_tree_bounding_box();
        }

        private void correct_tree_bounding_box()
        {
            // move everything so that the bottom left of the tree is at the origin (0,0);

            var bb = this.GetBoundingBoxOfTree();
            foreach (var n in this.Nodes)
            {
                n.Position = n.Position - bb.LowerLeft;
            }
        }

        public struct ParentChildConnection<U>
        {
            public readonly U Parent;
            public readonly U Child;

            public ParentChildConnection(U parent, U child)
            {
                this.Parent = parent;
                this.Child = child;
            }
        }

        public IEnumerable<ParentChildConnection<Node<T>>> EnumConnections()
        {
            foreach (var parent in this.Nodes)
            {
                foreach (var child in parent.EnumChildren())
                {
                    var connection = new ParentChildConnection<Node<T>>(parent, child);
                    yield return connection;
                }
            }
        }

        private static double _get_side(VisioAutomation.Geometry.Rectangle r, LayoutDirection direction)
        {
            switch (direction)
            {
                case (LayoutDirection.Up):
                    {
                        return r.Top;
                    }
                case (LayoutDirection.Down):
                    {
                        return r.Bottom;
                    }
                case (LayoutDirection.Left):
                    {
                        return r.Left;
                    }
                case (LayoutDirection.Right):
                    {
                        return r.Right;
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
            }
        }

        public static LayoutDirection GetOpposite(LayoutDirection direction)
        {
            switch (direction)
            {
                case (LayoutDirection.Up):
                    {
                        return LayoutDirection.Down;
                    }
                case (LayoutDirection.Down):
                    {
                        return LayoutDirection.Up;
                    }
                case (LayoutDirection.Left):
                    {
                        return LayoutDirection.Right;
                    }
                case (LayoutDirection.Right):
                    {
                        return LayoutDirection.Left;
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
            }
        }

        public Geometry.LineSegment GetConnectionLine(ParentChildConnection<Node<T>> connection)
        {
            var parent_rect = connection.Parent.Rect;
            var child_rect = connection.Child.Rect;

            double parent_x, parent_y;
            double child_x, child_y;

            if (TreeLayout<T>.IsVertical(this.Options.Direction))
            {
                parent_x = parent_rect.Center.X;
                child_x = child_rect.Center.X;

                parent_y = TreeLayout<T>._get_side(parent_rect, this.Options.Direction);
                child_y = TreeLayout<T>._get_side(child_rect, TreeLayout<T>.GetOpposite(this.Options.Direction));
            }
            else
            {
                var parent_dir = this.Options.Direction;
                var child_dir = TreeLayout<T>.GetOpposite(parent_dir);

                parent_x = TreeLayout<T>._get_side(parent_rect, parent_dir);
                child_x = TreeLayout<T>._get_side(child_rect, child_dir);

                parent_y = parent_rect.Center.Y;
                child_y = child_rect.Center.Y;
            }

            var parent_attach_point = new VisioAutomation.Geometry.Point(parent_x, parent_y);
            var child_attach_point = new VisioAutomation.Geometry.Point(child_x, child_y);

            return new Geometry.LineSegment(parent_attach_point, child_attach_point);
        }

        public static bool IsVertical(LayoutDirection direction)
        {
            return (direction == LayoutDirection.Up || direction == LayoutDirection.Down);
        }

        public VisioAutomation.Geometry.Point[] GetConnectionPolyline(ParentChildConnection<Node<T>> connection)
        {
            var lineseg = this.GetConnectionLine(connection);
            VisioAutomation.Geometry.Point m0, m1;

            var parent_attach_point = lineseg.Start;
            var child_attach_point = lineseg.End;
            var dif = lineseg.End - lineseg.Start;
            var a = (this.Options.LevelSeparation/2.0);
            var b = (this.Options.LevelSeparation/2.0);

            if (TreeLayout<T>.IsVertical(this.Options.Direction))
            {
                if (this.Options.Direction == LayoutDirection.Up)
                {
                    b = -b;
                }
                m0 = new VisioAutomation.Geometry.Point(lineseg.Start.X, lineseg.End.Y + b);
                m1 = new VisioAutomation.Geometry.Point(lineseg.End.X, lineseg.End.Y + b);
            }
            else
            {
                if (this.Options.Direction == LayoutDirection.Left)
                {
                    a = -a;
                }
                m0 = new VisioAutomation.Geometry.Point(lineseg.End.X - a, lineseg.Start.Y);
                m1 = new VisioAutomation.Geometry.Point(lineseg.End.X - a, lineseg.End.Y);
            }

            return new[] {lineseg.Start, m0, m1, lineseg.End};
        }

        public VisioAutomation.Geometry.Point[] GetConnectionBezier(ParentChildConnection<Node<T>> connection)
        {
            var lineseg = this.GetConnectionLine(connection);

            var parent_attach_point = lineseg.Start;
            var child_attach_point = lineseg.End;

            double scale = this.Options.LevelSeparation/2.0;
            var dif = child_attach_point.Subtract(parent_attach_point).Multiply(scale, scale);


            var handle_displacement = TreeLayout<T>.IsVertical(this.Options.Direction)
                                          ? new VisioAutomation.Geometry.Point(0, dif.Y)
                                          : new VisioAutomation.Geometry.Point(dif.X, 0);

            var h1 = parent_attach_point.Add(handle_displacement);
            var h2 = child_attach_point.Add( handle_displacement.Multiply(-1, -1));

            return new[] {parent_attach_point, h1, h2, child_attach_point};
        }

        public static Node<T> CreateLayoutTree<TA>(
            TA root,
            System.Func<TA, IEnumerable<TA>> enum_children,
            System.Func<TA, T> func_get_data,
            System.Func<TA, VisioAutomation.Geometry.Size> func_get_size)
        {
            var walkevents = GenTreeOps.Algorithms.Walk<TA>(root, n => enum_children(n));
            return TreeLayout<T>._create_layout_tree(walkevents, func_get_data, func_get_size);
        }

        private static Node<T> _create_layout_tree<TA>(
            IEnumerable<GenTreeOps.WalkEvent<TA>> walkevents,
            System.Func<TA, T> func_get_data,
            System.Func<TA, VisioAutomation.Geometry.Size> func_get_size)
        {
            var stack = new Stack<Node<T>>();
            Node<T> layout_root = null;
            foreach (var walkevent in walkevents)
            {
                if (walkevent.Type == GenTreeOps.WalkEventType.EventEnter)
                {
                    Node<T> parent = null;
                    if (stack.Count > 0)
                    {
                        parent = stack.Peek();
                    }

                    var data = func_get_data(walkevent.Node);
                    var size = func_get_size(walkevent.Node);

                    var layout_node = new Node<T>(size, data);
                    if (parent != null)
                    {
                        parent.AddChild(layout_node);
                    }
                    stack.Push(layout_node);

                    if (layout_root == null)
                    {
                        layout_root = layout_node;
                    }
                }
                else if (walkevent.Type == GenTreeOps.WalkEventType.EventExit)
                {
                    var layout_node = stack.Pop();
                }
            }

            return layout_root;
        }
    }
}