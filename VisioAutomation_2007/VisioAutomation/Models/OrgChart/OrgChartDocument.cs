using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Models.OrgChart
{
    public class OrgChartDocument
    {
        public List<Node> OrgCharts { get; private set; }

        public LayoutOptions LayoutOptions;

        public OrgChartDocument()
        {
            this.OrgCharts = new List<Node>();
            this.LayoutOptions = new LayoutOptions();
        }

        private VA.Models.InternalTree.Node<object> node_to_layout_node(Node n)
        {
            var nodesize = n.Size.GetValueOrDefault(this.LayoutOptions.DefaultNodeSize);
            var newnode = new VA.Models.InternalTree.Node<object>(nodesize, n);
            return newnode;
        }

        public void Render(IVisio.Application app)
        {
            var orgchartdrawing = this;

            if (orgchartdrawing == null)
            {
                throw new System.ArgumentNullException("orgchartdrawing");
            }

            if (app == null)
            {
                throw new System.ArgumentNullException("app");
            }

            if (orgchartdrawing.OrgCharts.Count < 1)
            {
                throw new System.ArgumentException("orgchart must have at least one root");
            }

            foreach (var root in orgchartdrawing.OrgCharts)
            {
                if (root == null)
                {
                    throw new System.ArgumentException("Org chart has root node set to null", "orgchartdrawing");
                }
            }

            int majorver = int.Parse(app.Version.Split('.')[0]);
            bool is_visio_2013 = majorver >= 15;

            const string orgchart_vst = "orgch_u.vst";
            string orgchart_master_node_name = is_visio_2013 ? "Position Belt" : "Position";
            const string dyncon_master_name = "Dynamic connector";
            const double border_width = 0.5;

            var doc_node = new VA.DOM.Document(orgchart_vst, IVisio.VisMeasurementSystem.visMSUS);

            var trees = new List<IList<VisioAutomation.Models.InternalTree.Node<object>>>();

            foreach (var root in orgchartdrawing.OrgCharts)
            {
                // Construct a layout tree from the hierarchy
                var treenodes = VA.Internal.TreeOps.CopyTree(
                    orgchartdrawing.OrgCharts[0],
                    n => n.Children,
                    n => node_to_layout_node(n),
                    (p, c) => p.AddChild(c));

                trees.Add(treenodes);

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

                // vis.ActiveWindow.ShowConnectPoints = 0;

                var page_node = new VA.DOM.Page();
                doc_node.Pages.Add(page_node);

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
                    .Select(centerpoint => page_node.Shapes.Drop(orgchart_master_node_name, null, centerpoint))
                    .ToList();


                // For each OrgChart object, attach the shape that corresponds to it
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    var orgnode = (Node)treenodes[i].Data;
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
                            var connector = page_node.Shapes.Connect(dyncon_master_name, null, parent_shape, child_shape);
                        }
                    }
                }
                else
                {
                    foreach (var connection in layout.EnumConnections())
                    {
                        var bez = layout.GetConnectionBezier(connection);
                        page_node.Shapes.DrawBezier(bez);
                    }
                }

                // Set the Text Labels on each Org node
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    var orgnode = (Node)treenodes[i].Data;
                    var shape = (VA.DOM.BaseShape)orgnode.DOMNode;
                    shape.Text = new VA.Text.Markup.TextElement(orgnode.Text);
                }

                var page_size_with_border = bb.Size.Add(border_width * 2, border_width * 2.0);
                page_node.Size = page_size_with_border;
                page_node.ResizeToFit = true;
                page_node.ResizeToFitMargin = new VA.Drawing.Size(border_width * 2, border_width * 2.0);
            } // finish handling root node

            var doc = doc_node.Render(app);

            foreach (var treenodes in trees)
            {
                var orgnodes = treenodes.Select(i => i.Data).Cast<Node>();
                var orgnodes_with_urls = orgnodes.Where(n => n.URL != null);
                var all_urls = orgnodes_with_urls.Select(n => new { orgnode = n, shape = (VA.DOM.BaseShape)n.DOMNode, url = n.URL.Trim() });

                foreach (var url in all_urls)
                {
                    var hlink = url.orgnode.VisioShape.Hyperlinks.Add();
                    hlink.Name = "Row_1";
                    hlink.Address = url.orgnode.URL;
                }

                // Attach all the orgchart nodes to the Visio shapes that were created
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    var orgnode = (Node)treenodes[i].Data;
                    var shape = (VA.DOM.BaseShape)orgnode.DOMNode;
                    orgnode.VisioShape = shape.VisioShape;
                }
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