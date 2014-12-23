using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Tree
{
    public class TreeLayout
    {
        const string basic_stencil_name = "basic_u.vss";
        const string connectors_stencil_name = "connec_u.vss";
        string rect_master_name = "Rectangle";
        private string dc_master_name = "Dynamic Connector";

        public LayoutOptions LayoutOptions { get; set; }

        public TreeLayout()
        {
            this.LayoutOptions = new LayoutOptions();
        }

        private VA.Models.InternalTree.Node<object> node_to_layout_node(Node n)
        {
            var nodesize = n.Size.GetValueOrDefault(this.LayoutOptions.DefaultNodeSize);
            var newnode = new VA.Models.InternalTree.Node<object>(nodesize, n);
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
            var treenodes = VA.Internal.TreeOps.CopyTree(
                drawing.Root,
                n => n.Children,
                n => node_to_layout_node(n),
                (p, c) => p.AddChild(c));

            // Perform the layout
            var layout = new VA.Models.InternalTree.TreeLayout<object>();

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
            var basic_stencil = documents.OpenStencil(basic_stencil_name);
            var connectors_stencil = documents.OpenStencil(connectors_stencil_name);
            var basic_masters = basic_stencil.Masters;
            var connectors_masters = connectors_stencil.Masters;

            var node_master = basic_masters[rect_master_name];
            var connector_master = connectors_masters[dc_master_name];

            var page_node = new VA.DOM.Page();

            var page_size = bb.Size.Add(border_width*2, border_width*2.0);

            // fixup the nodes so that they render on the page
            foreach (var i in treenodes)
            {
                i.Position = i.Position.Add(border_width, border_width);
            }

            var centerpoints = treenodes.Select(tn => tn.Rect.Center).ToList();
            var master_nodes = centerpoints.Select(cp => page_node.Shapes.Drop(node_master, cp)).ToList();

            // For each OrgChart object, attach the shape that corresponds to it
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var tree_node = (VA.Models.Tree.Node)treenodes[i].Data;
                DOM.Shape master_node = master_nodes[i];
                tree_node.DOMNode = master_node;

                if (tree_node.Cells!=null)
                {
                    master_node.Cells = tree_node.Cells.ShallowCopy();
                }

                master_node.Cells.Width = treenodes[i].Size.Width;
                master_node.Cells.Height = treenodes[i].Size.Height;
                master_node.Text = tree_node.Text;
            }

            if (this.LayoutOptions.ConnectorType  == ConnectorType.DynamicConnector)
            {
                var orgchart_nodes = treenodes.Select(tn => tn.Data).Cast<Node>();

                foreach (var parent in orgchart_nodes)
                {
                    foreach (var child in parent.Children)
                    {
                        var parent_shape = (VA.DOM.BaseShape)parent.DOMNode;
                        var child_shape = (VA.DOM.BaseShape)child.DOMNode;
                        var connector = page_node.Shapes.Connect(connector_master, parent_shape, child_shape);
                        connector.Cells = this.LayoutOptions.ConnectorCells;
                    }
                }
            }
            else if  (this.LayoutOptions.ConnectorType == ConnectorType.CurvedBezier)
            {
                foreach (var connection in layout.EnumConnections())
                {
                    var bez = layout.GetConnectionBezier(connection);
                    var shape = page_node.Shapes.DrawBezier(bez);
                    shape.Cells = this.LayoutOptions.ConnectorCells;
                }
            }
            else if (this.LayoutOptions.ConnectorType == ConnectorType.PolyLine)
            {
                foreach (var connection in layout.EnumConnections())
                {
                    var polyline = layout.GetConnectionPolyline(connection);
                    var shape = page_node.Shapes.DrawPolyLine(polyline);
                    shape.Cells = this.LayoutOptions.ConnectorCells;
                }
            }
            else
            {
                string msg = "Unsupported Connector Type";
                throw new VA.AutomationException(msg);
            }

            page_node.Size = page_size;
            page_node.Render(page);

            // Attach all the orgchart nodes to the Visio shapes that were created
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var orgnode = (Node) treenodes[i].Data;
                var shape = (VA.DOM.BaseShape)orgnode.DOMNode;
                orgnode.VisioShape = shape.VisioShape;
            }
        }

        private VA.Models.InternalTree.LayoutDirection map_direction2(LayoutDirection input_dir)
        {
            VA.Models.InternalTree.LayoutDirection dir;
            if (input_dir == LayoutDirection.Down)
            {
                dir = VA.Models.InternalTree.LayoutDirection.Down;
            }
            else if (input_dir == LayoutDirection.Up)
            {
                dir = VA.Models.InternalTree.LayoutDirection.Up;
            }
            else if (input_dir == LayoutDirection.Left)
            {
                dir = VA.Models.InternalTree.LayoutDirection.Left;
            }
            else if (input_dir == LayoutDirection.Right)
            {
                dir = VA.Models.InternalTree.LayoutDirection.Right;
            }
            else
            {
                dir = VA.Models.InternalTree.LayoutDirection.Down;
            }
            return dir;
        }
    }
}