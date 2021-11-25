﻿using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.Models.Layouts.InternalTree;

namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class OrgChartDocument
    {
        public List<Node> OrgCharts { get; }

        public OrgChartLayoutOptions OrgChartLayoutOptions;
        public OrgChartStyling Styling = new OrgChartStyling();

        public OrgChartDocument()
        {
            this.OrgCharts = new List<Node>();
            this.OrgChartLayoutOptions = new OrgChartLayoutOptions();
        }

        private Node<object> node_to_layout_node(Node n)
        {
            var nodesize = n.Size.GetValueOrDefault(this.OrgChartLayoutOptions.DefaultNodeSize);
            var newnode = new Node<object>(nodesize, n);
            return newnode;
        }

        public void Render(IVisio.Application app)
        {
            var orgchartdrawing = this;

            if (orgchartdrawing == null)
            {
                throw new System.ArgumentNullException(nameof(orgchartdrawing));
            }

            if (app == null)
            {
                throw new System.ArgumentNullException(nameof(app));
            }

            if (orgchartdrawing.OrgCharts.Count < 1)
            {
                throw new System.ArgumentException("orgchart must have at least one root");
            }

            foreach (var root in orgchartdrawing.OrgCharts)
            {
                if (root == null)
                {
                    throw new System.ArgumentException("Org chart has root node set to null");
                }
            }


            var ver = Application.ApplicationHelper.GetVersion(app);
            int majorver = ver.Major;
            bool is_visio_2013_or_newer = majorver >= 15;
            
            string orgchart_template = is_visio_2013_or_newer ? this.Styling.Visio2013Template : this.Styling.Visio2010Template;
            string orgchart_node_master_name = is_visio_2013_or_newer ? this.Styling.Visio2013NodeMaster : this.Styling.Visio2010NodeMaster;
            string orgchart_dyncon_master_name = is_visio_2013_or_newer ? this.Styling.Visio2013ConnectorMaster : this.Styling.Visio2010ConnectorMaster;


            var doc_node = new Dom.Document(orgchart_template, IVisio.VisMeasurementSystem.visMSUS);

            var trees = new List<IList<Node<object>>>();

            foreach (var root in orgchartdrawing.OrgCharts)
            {
                // Construct a layout tree from the hierarchy
                var treenodes = GenTreeOps.Algorithms.CopyTree(
                    orgchartdrawing.OrgCharts[0],
                    n => n.Children,
                    n => this.node_to_layout_node(n),
                    (p, c) => p.AddChild(c));

                trees.Add(treenodes);

                // Perform the layout
                var layout = new TreeLayout<object>();

                layout.Options.Direction = this.map_direction2(this.OrgChartLayoutOptions.Direction);
                layout.Options.LevelSeparation = 1;
                layout.Options.SiblingSeparation = 0.25;
                layout.Options.SubtreeSeparation = 1;

                layout.Root.AddChild(treenodes[0]);
                layout.PerformLayout();

                // Render the Document in Visio
                var bb = layout.GetBoundingBoxOfTree();

                // vis.ActiveWindow.ShowConnectPoints = 0;

                var page_node = new Dom.Page();
                doc_node.Pages.Add(page_node);

                // fixup the nodes so that they render on the page
                foreach (var i in treenodes)
                {
                    i.Position = i.Position.Add(this.OrgChartLayoutOptions.PageBorderWidth, this.OrgChartLayoutOptions.PageBorderWidth);
                }

                var centerpoints = new VisioAutomation.Geometry.Point[treenodes.Count];
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    centerpoints[i] = treenodes[i].Rect.Center;
                }

                // TODO: Add support for Left to right , Right to Left, and Bottom to Top Layouts

                var vmasters = centerpoints
                    .Select(centerpoint => page_node.Shapes.Drop(orgchart_node_master_name, null, centerpoint))
                    .ToList();


                // For each OrgChart object, attach the shape that corresponds to it
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    var orgnode = (Node)treenodes[i].Data;
                    orgnode.DomNode = vmasters[i];
                    vmasters[i].Cells.XFormWidth = treenodes[i].Size.Width;
                    vmasters[i].Cells.XFormHeight = treenodes[i].Size.Height;
                }

                if (this.OrgChartLayoutOptions.UseDynamicConnectors)
                {
                    var orgchart_nodes = treenodes.Select(tn => tn.Data).Cast<Node>();

                    foreach (var parent in orgchart_nodes)
                    {
                        foreach (var child in parent.Children)
                        {
                            var parent_shape = (Dom.BaseShape)parent.DomNode;
                            var child_shape = (Dom.BaseShape)child.DomNode;
                            var connector = page_node.Shapes.Connect(orgchart_dyncon_master_name, null, parent_shape, child_shape);
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
                    var shape = (Dom.BaseShape)orgnode.DomNode;
                    shape.Text = new VisioAutomation.Models.Text.Element(orgnode.Text);
                }

                var page_size_with_border = bb.Size.Add(this.OrgChartLayoutOptions.PageBorderWidth * 2, this.OrgChartLayoutOptions.PageBorderWidth * 2.0);
                page_node.Size = page_size_with_border;
                page_node.ResizeToFit = true;
                page_node.ResizeToFitMargin = new VisioAutomation.Geometry.Size(this.OrgChartLayoutOptions.PageBorderWidth * 2, this.OrgChartLayoutOptions.PageBorderWidth * 2.0);
            } // finish handling root node

            var doc = doc_node.Render(app);

            foreach (var treenodes in trees)
            {
                var orgnodes = treenodes.Select(i => i.Data).Cast<Node>();
                var orgnodes_with_urls = orgnodes.Where(n => n.Url != null);
                var all_urls = orgnodes_with_urls.Select(n => new { orgnode = n, shape = (Dom.BaseShape)n.DomNode, url = n.Url.Trim() });

                foreach (var url in all_urls)
                {
                    var hlink = url.orgnode.VisioShape.Hyperlinks.Add();
                    hlink.Name = "Row_1";
                    hlink.Address = url.orgnode.Url;
                }

                // Attach all the orgchart nodes to the Visio shapes that were created
                foreach (int i in Enumerable.Range(0, treenodes.Count))
                {
                    var orgnode = (Node)treenodes[i].Data;
                    var shape = (Dom.BaseShape)orgnode.DomNode;
                    orgnode.VisioShape = shape.VisioShape;
                }
            }
        }

        private Layouts.InternalTree.LayoutDirection map_direction2(OrgChartLayoutDirection input_dir)
        {
            Layouts.InternalTree.LayoutDirection dir;
            if (input_dir == OrgChartLayoutDirection.Down)
            {
                dir = Layouts.InternalTree.LayoutDirection.Down;
            }
            else if (input_dir == OrgChartLayoutDirection.Up)
            {
                dir = Layouts.InternalTree.LayoutDirection.Up;
            }
            else if (input_dir == OrgChartLayoutDirection.Left)
            {
                dir = Layouts.InternalTree.LayoutDirection.Left;
            }
            else if (input_dir == OrgChartLayoutDirection.Right)
            {
                dir = Layouts.InternalTree.LayoutDirection.Right;
            }
            else
            {
                dir = Layouts.InternalTree.LayoutDirection.Down;
            }
            return dir;
        }

    }
}