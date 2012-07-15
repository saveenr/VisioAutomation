using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAL = VisioAutomation.Layout;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Models.OrgChart

{
    public class OrgChartLayout
    {
        public LayoutOptions LayoutOptions;

        public OrgChartLayout()
        {
            this.LayoutOptions = new LayoutOptions();
        }

        private VA.Layout.Models.InternalTree.Node<object> node_to_layout_node(Node n)
        {
            var nodesize = n.Size.GetValueOrDefault(this.LayoutOptions.DefaultNodeSize);
            var newnode = new VA.Layout.Models.InternalTree.Node<object>(nodesize, n);
            return newnode;
        }

        internal void RenderToVisio(Drawing orgchartdrawing, IVisio.Application app)
        {
            if (orgchartdrawing == null)
            {
                throw new System.ArgumentNullException("orgchartdrawing");
            }

            if (app == null)
            {
                throw new System.ArgumentNullException("app");
            }

            if (orgchartdrawing.Root == null)
            {
                throw new System.ArgumentException("Org chart has root node set to null", "orgchartdrawing");
            }

            const string xorgchart_vst = "orgch_u.vst";
            const string xorgchart_vss = "orgch_u.vss";
            string xorgchart_master_node_name = "Position";
            const double border_width = 0.5;


            // Construct a layout tree from the hierarchy
            var treenodes = VA.Internal.TreeUtil.CopyTree(
                orgchartdrawing.Root,
                n => n.Children,
                n => node_to_layout_node(n),
                (p, c) => p.AddChild(c));

            // Perform the layout
            var layout = new VA.Layout.Models.InternalTree.TreeLayout<object>();


           layout.Options.Direction = map_direction2(this.LayoutOptions.Direction);
            layout.Options.LevelSeparation = 1;
            layout.Options.SiblingSeparation = 0.25;
            layout.Options.SubtreeSeparation = 1;

            layout.Root.AddChild(treenodes[0]);
            layout.PerformLayout();

            // Render the Document in Visio
            var bb = layout.GetBoundingBoxOfTree();

            // vis.ActiveWindow.ShowConnectPoints = 0;
            var documents = app.Documents;
            var stencil = documents.OpenStencil(xorgchart_vss);
            var master = stencil.Masters[xorgchart_master_node_name];
            var dc_master = stencil.Masters["Dynamic Connector"];
            var doc = documents.AddEx(xorgchart_vst, IVisio.VisMeasurementSystem.visMSUS, 0, 0);

            var domshapescol = new VA.DOM.ShapeList();
            
            // fixup the nodes so that they render on the page
            foreach (var i in treenodes)
            {
                i.Position = i.Position.Add(border_width, border_width);
            }

            var centerpoints = new VA.Drawing.Point[treenodes.Count];
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                centerpoints[i] = treenodes[i].Rect.Center;
            }

            // TODO: Add support for Left to right , Right to Left, and Bottom to Top Layouts


            var vmasters = centerpoints
                .Select(centerpoint => domshapescol.Drop(master, centerpoint))
                .ToList();


            // For each OrgChart object, attach the shape that corresponds to it
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var orgnode = (Node) treenodes[i].Data;
                orgnode.DOMNode = vmasters[i];
                vmasters[i].Cells.Width = treenodes[i].Size.Width;
                vmasters[i].Cells.Height = treenodes[i].Size.Height;
            }

            if (this.LayoutOptions.UseDynamicConnectors)
            {
                var orgchart_nodes = treenodes.Select(tn => tn.Data).Cast<Node>();

                foreach (var parent in orgchart_nodes)
                {
                    foreach (var child in parent.Children)
                    {
                        var parent_shape = (VA.DOM.BaseShape)parent.DOMNode;
                        var child_shape = (VA.DOM.BaseShape)child.DOMNode;
                        var connector = domshapescol.Connect(dc_master,parent_shape, child_shape);
                    }
                }
            }
            else
            {
                foreach (var connection in layout.EnumConnections())
                {
                    var bez = layout.GetConnectionBezier(connection);
                    domshapescol.DrawBezier(bez);
                }
            }

            // Set the Text Labels on each Org node
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var orgnode = (Node) treenodes[i].Data;
                var shape = (VA.DOM.BaseShape)orgnode.DOMNode;
                shape.Text = new VA.Text.Markup.TextElement(orgnode.Text);
            }

            var page = doc.Pages.Add();
            var page_size_with_border = bb.Size.Add(border_width*2, border_width*2.0);
            page.SetSize(page_size_with_border);

            domshapescol.Render(page);

            var orgnodes = treenodes.Select(i => i.Data).Cast<Node>();
            var orgnodes_with_urls = orgnodes.Where(n => n.URL != null);
            var all_urls = orgnodes_with_urls.Select( n=>  new { orgnode = n, shape = (VA.DOM.BaseShape) n.DOMNode, url = n.URL.Trim() } );

            foreach (var url in all_urls)
            {
                var hlink = url.orgnode.VisioShape.Hyperlinks.Add();
                hlink.Name = "Row_1";
                hlink.Address = url.orgnode.URL;
            }
            
            // Attach all the orgchart nodes to the Visio shapes that were created
            foreach (int i in Enumerable.Range(0, treenodes.Count))
            {
                var orgnode = (Node) treenodes[i].Data;
                var shape = (VA.DOM.BaseShape)orgnode.DOMNode;
                orgnode.VisioShape = shape.VisioShape;
            }
        }

        private VA.Layout.Models.InternalTree.LayoutDirection map_direction2(LayoutDirection input_dir)
        {
            VA.Layout.Models.InternalTree.LayoutDirection dir;
            if (input_dir == LayoutDirection.Down)
            {
                dir = VA.Layout.Models.InternalTree.LayoutDirection.Down;
            }
            else if (input_dir == LayoutDirection.Up)
            {
                dir = VA.Layout.Models.InternalTree.LayoutDirection.Up;
            }
            else if (input_dir == LayoutDirection.Left)
            {
                dir = VA.Layout.Models.InternalTree.LayoutDirection.Left;
            }
            else if (input_dir == LayoutDirection.Right)
            {
                dir = VA.Layout.Models.InternalTree.LayoutDirection.Right;
            }
            else
            {
                dir = VA.Layout.Models.InternalTree.LayoutDirection.Down;
            }
            return dir;
        }
    }
}