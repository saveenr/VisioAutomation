using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Tree
{
    public class TreeLayout
    {
        const string stencil_name = "basic_u.vss";
        string master_name = "Rectangle";
        private string dc_name = "Dynamic Connector";

        public LayoutOptions LayoutOptions { get; set; }

        public TreeLayout()
        {
            this.LayoutOptions = new LayoutOptions();
        }

        private VA.Layout.InternalTree.Node<object> node_to_layout_node(Node n)
        {
            var nodesize = n.Size.GetValueOrDefault(this.LayoutOptions.DefaultNodeSize);
            var newnode = new VA.Layout.InternalTree.Node<object>(nodesize, n);
            return newnode;
        }

        internal void RenderToVisio(Drawing drawing, IVisio.Page page)
        {
            if (drawing == null)
            {
                throw new System.ArgumentNullException("drawing");
            }

            if (page== null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (drawing.Root == null)
            {
                throw new System.ArgumentException("Tree has root node set to null", "drawing");
            }

            const double border_width = 0.5;

            // Construct a layout tree from the hierarchy
            var treenodes = VA.Internal.TreeUtil.CopyTree(
                drawing.Root,
                n => n.Children,
                n => node_to_layout_node(n),
                (p, c) => p.AddChild(c));

            // Perform the layout
            var layout = new VA.Layout.InternalTree.TreeLayout<object>();

            layout.Options.Direction = map_direction2(this.LayoutOptions.Direction);
            layout.Options.LevelSeparation = 1;
            layout.Options.SiblingSeparation = 0.25;
            layout.Options.SubtreeSeparation = 1;

            layout.Root.AddChild(treenodes[0]);
            layout.PerformLayout();

            // Render the Document in Visio
            var bb = layout.GetBoundingBoxOfTree();

            var app = page.Application;
            var documents = app.Documents;
            var stencil = documents.OpenStencil(stencil_name);
            var masters = stencil.Masters;
            var node_master = masters[master_name];
            var connector_master = masters[dc_name];

            var dom_doc = new VA.DOM.Document();
            dom_doc.ResolveVisioShapeObjects = true;

            var page_size = bb.Size.Add(border_width*2, border_width*2.0);

            // fixup the nodes so that they render on the page
            foreach (var i in treenodes)
            {
                i.Position = i.Position.Add(border_width, border_width);
            }

            // TODO: Add support for Left to right , Right to Left, and Bottom to Top Layouts
            var centerpoints = treenodes.Select(tn => tn.Rect.Center).ToList();
            var dom_masters = centerpoints.Select(cp => dom_doc.Drop(node_master, cp)).ToList();

            // For each OrgChart object, attach the shape that corresponds to it
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var tree_node = (VA.Layout.Tree.Node) treenodes[i].Data;
                DOM.Master dom_master = dom_masters[i];
                tree_node.DOMNode = dom_master;

                if (tree_node.ShapeCells!=null)
                {
                    dom_master.ShapeCells = tree_node.ShapeCells.ShallowCopy();
                }

                dom_master.ShapeCells.Width = treenodes[i].Size.Width;
                dom_master.ShapeCells.Height = treenodes[i].Size.Height;
                dom_master.Text = tree_node.Text;
            }

            if (this.LayoutOptions.UseDynamicConnectors)
            {
                var orgchart_nodes = treenodes.Select(tn => tn.Data).Cast<Node>();

                foreach (var parent in orgchart_nodes)
                {
                    foreach (var child in parent.Children)
                    {
                        var parent_shape = (VA.DOM.Shape)parent.DOMNode;
                        var child_shape = (VA.DOM.Shape)child.DOMNode;
                        var connector = dom_doc.Connect(connector_master, parent_shape, child_shape);
                    }
                }
            }
            else
            {
                foreach (var connection in layout.EnumConnections())
                {
                    var bez = layout.GetConnectionBezier(connection);
                    var shape = dom_doc.DrawBezier(bez);
                }
            }

            dom_doc.Render(page);

            page.SetSize(page_size);

            // Attach all the orgchart nodes to the Visio shapes that were created
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var orgnode = (Node) treenodes[i].Data;
                var shape = (VA.DOM.Shape)orgnode.DOMNode;
                orgnode.VisioShape = shape.VisioShape;
            }
        }

        private VA.Layout.InternalTree.LayoutDirection map_direction2(LayoutDirection input_dir)
        {
            VA.Layout.InternalTree.LayoutDirection dir;
            if (input_dir == LayoutDirection.Down)
            {
                dir = VA.Layout.InternalTree.LayoutDirection.Down;
            }
            else if (input_dir == LayoutDirection.Up)
            {
                dir = VA.Layout.InternalTree.LayoutDirection.Up;
            }
            else if (input_dir == LayoutDirection.Left)
            {
                dir = VA.Layout.InternalTree.LayoutDirection.Left;
            }
            else if (input_dir == LayoutDirection.Right)
            {
                dir = VA.Layout.InternalTree.LayoutDirection.Right;
            }
            else
            {
                dir = VA.Layout.InternalTree.LayoutDirection.Down;
            }
            return dir;
        }
    }
}