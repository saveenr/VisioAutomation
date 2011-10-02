using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    public class BoxLayout<T>
    {
        public LayoutOptions LayoutOptions;

        private Node<T> _root;

        public BoxLayout() :
            this(LayoutDirection.Vertical)
        {
        }

        public BoxLayout(LayoutDirection dir)
        {
            this.LayoutOptions = new LayoutOptions(); 
            this._root = new Node<T>(dir);
        }

        public Node<T> Root
        {
            get { return _root; }
            set { _root = value; }
        }

        public void PerformLayout()
        {
            if (this.Root.ChildCount < 1)
            {
                throw new VA.AutomationException("Root must contain at least one child");
            }

            // The first stage is to figure out how big the boxes need to be
            this.CalculateSizes();

            // having that, we then use the layout options to put them in the correct positions
            this.Place(this.LayoutOptions.Origin);

            // Place doesn't calculate the reserved rectangle of the root node
            // so we do it here. because the root contains "everything" the Reserved Rectangle
            // is the same is its rectangle calculated by Place
            this.Root.ReservedRectangle = this.Root.Rectangle;
        }

        private void CalculateSizes()
        {
            // this method calculates the sizes of nodes
            _CalculateSizeNode(_root);
        }

        private void _CalculateSizeNode(Node<T> node)
        {
            //calculate the size of the children
            foreach (var child_el in node.Children)
            {
                _CalculateSizeNode(child_el);
            }

            double child_height_sum = 0;
            double child_width_max = 0;
            double child_height_max = 0;
            double child_width_sum = 0;
            double h = node.Height.GetValueOrDefault(LayoutOptions.DefaultHeight);
            double w = node.Width.GetValueOrDefault(LayoutOptions.DefaultWidth);

            double padx = node.Padding;
            double pady = node.Padding;

            foreach (var child_el in node.Children)
            {
                child_height_sum += child_el.Height.Value;
                child_height_max = System.Math.Max(child_height_max, child_el.Height.Value);
                child_width_sum += child_el.Width.Value;
                child_width_max = System.Math.Max(child_width_max, child_el.Width.Value);
            }

            // Account for child separation
            int num_seps = System.Math.Max(0, node.ChildCount - 1);
            double total_sepy = (node.Direction == LayoutDirection.Vertical) ? num_seps*node.ChildSeparation : 0.0;
            double total_sepx = (node.Direction == LayoutDirection.Horizonal) ? num_seps*node.ChildSeparation : 0.0;

            child_height_sum += total_sepy;
            child_width_sum += total_sepx;

            if (node.Direction == LayoutDirection.Vertical)
            {
                node.Height = System.Math.Max(h, child_height_sum);
                node.Width = System.Math.Max(w, child_width_max);
            }
            else if (node.Direction == LayoutDirection.Horizonal)
            {
                node.Height = System.Math.Max(h, child_height_max);
                node.Width = System.Math.Max(w, child_width_sum);
            }

            node.Height = node.Height.Value + (2*pady);
            node.Width = node.Width.Value + (2*padx);
        }

        private void Place(VA.Drawing.Point origin)
        {
            // this method calculates the positions on nodes
            _PlaceNode(_root, origin);
        }

        private void _PlaceNode(Node<T> node, VA.Drawing.Point origin)
        {
            if (node == null)
            {
                throw new System.ArgumentNullException("node");
            }

            double sign_x = (LayoutOptions.DirectionHorizontal == VA.Layout.BoxLayout.DirectionHorizontal.LeftToRight) ? 1.0 : -1.0;
            double sign_y = (LayoutOptions.DirectionVertical == VA.Layout.BoxLayout.DirectionVertical.BottomToTop) ? 1.0 : -1.0;

            // Calculate the final rectangle to place the current node

            double miny = (LayoutOptions.DirectionVertical == VA.Layout.BoxLayout.DirectionVertical.TopToBottom)
                              ? origin.Y - node.Height.Value
                              : origin.Y;

            double minx = (LayoutOptions.DirectionHorizontal == VA.Layout.BoxLayout.DirectionHorizontal.LeftToRight)
                              ? origin.X
                              : origin.X - node.Width.Value;

            double maxx = minx + node.Width.Value;

            double maxy = miny + node.Height.Value;

            var rect = new VA.Drawing.Rectangle(minx, miny, maxx, maxy);
            node.Rectangle = rect;

            var current_point = origin;
            double pad_x = node.Padding;
            double pad_y = node.Padding;

            foreach (var cur_el in node.Children)
            {
                // Calculate where the child will be placed, taking into account the direction and alignment
                var child_origin = current_point;

                var reserved_width = node.Direction == LayoutDirection.Vertical ? node.Width.Value - 2*node.Padding: cur_el.Width.Value;
                var reserved_height = node.Direction == LayoutDirection.Horizonal? node.Height.Value - 2*node.Padding: cur_el.Height.Value;
                var reserved_size = new VA.Drawing.Size(reserved_width, reserved_height);
                cur_el.ReservedRectangle = new VA.Drawing.Rectangle(child_origin.Add(pad_x,pad_y),reserved_size);

                if (node.Direction == LayoutDirection.Vertical)
                {
                    var halign = cur_el.AlignmentHorizontal;

                    double delta_width = node.Width.Value - (2*pad_x) - cur_el.Width.Value;
                    double align_delta_x = (halign == VA.Drawing.AlignmentHorizontal.Left) ? 0.0 : delta_width;
                    double align_factor_x = (halign == VA.Drawing.AlignmentHorizontal.Center) ? 0.5 : 1.0;

                    child_origin = current_point.Add(sign_x*align_factor_x*align_delta_x, 0);
                }
                else
                {
                    var valign = cur_el.AlignmentVertical;

                    double delta_height = node.Height.Value - (2*pad_y) - cur_el.Height.Value;
                    double align_delta_y = (valign == VA.Drawing.AlignmentVertical.Bottom) ? 0.0 : delta_height;
                    double align_factor_y = (valign == VA.Drawing.AlignmentVertical.Center) ? 0.5 : 1.0;
                    child_origin = current_point.Add(0, sign_y*align_factor_y*align_delta_y);
                }

                child_origin = child_origin.Add(sign_x*pad_x, sign_y*pad_y);

                // render the child
                _PlaceNode(cur_el, child_origin);

                // move to the next place to start placing a child
                if (node.Direction == LayoutDirection.Vertical)
                {
                    current_point = current_point.Add(0, sign_y*cur_el.Height.Value);
                    current_point = current_point.Add(0, sign_y*node.ChildSeparation);
                }
                else if (node.Direction == LayoutDirection.Horizonal)
                {
                    current_point = current_point.Add(sign_x*cur_el.Width.Value, 0);
                    current_point = current_point.Add(sign_x*node.ChildSeparation, 0);
                }
            }
        }

        public IEnumerable<Node<T>> Nodes
        {
            get { return VA.Internal.TreeTraversal.PreOrder(this.Root, n => n.Children); }
        }
    }
}