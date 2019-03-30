using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using MSAGL = Microsoft.Msagl;
using VA = VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    class MsaglRenderer
    {
        private string _edge_master_name = "Dynamic Connector";
        private string _edge_stencil_name = "connec_u.vss";

        private VA.Geometry.Rectangle _mg_bb;
        private VA.Geometry.Rectangle _layout_bb;

        public Dom.ShapeCells DefaultBezierConnectorShapeCells { get; set; }
        public VA.Geometry.Size DefaultBezierConnectorLabelBoxSize { get; set; }
        public MsaglLayoutOptions LayoutOptions { get; set; }

        private double ScaleToMsagl
        {
            get { return this.LayoutOptions.ScalingFactor; }
        }

        private double ScaleToDocument
        {
            get { return 1.0 / this.LayoutOptions.ScalingFactor; }
        }

        public MsaglRenderer()
        {
            this.LayoutOptions = new MsaglLayoutOptions();

            this.DefaultBezierConnectorShapeCells = new Dom.ShapeCells();
            this.DefaultBezierConnectorShapeCells.LinePattern = 0;
            this.DefaultBezierConnectorShapeCells.LineWeight = 0.0;
            this.DefaultBezierConnectorShapeCells.FillPattern = 0;
            this.DefaultBezierConnectorLabelBoxSize = new VA.Geometry.Size(1.0, 0.5);
        }

        private VA.Geometry.Point _to_document_coordinates(VA.Geometry.Point point)
        {
            var np = point.Add(-this._mg_bb.Left, -this._mg_bb.Bottom).Multiply(this.ScaleToDocument, this.ScaleToDocument);
            return np;
        }

        private VA.Geometry.Rectangle _to_document_coordinates(VA.Geometry.Rectangle rect)
        {
            var nr = rect.Add(-this._mg_bb.Left, -this._mg_bb.Bottom).Multiply(this.ScaleToDocument, this.ScaleToDocument);
            return nr;
        }

        private VA.Geometry.Size _to_mg_coordinates(VA.Geometry.Size s)
        {
            return s.Multiply(this.ScaleToMsagl, this.ScaleToMsagl);
        }

        private void validate_connectors(DirectedGraphLayout layout_diagram)
        {
            foreach (var layout_connector in layout_diagram.Connectors)
            {
                if (layout_connector.ID == null)
                {
                    throw new System.ArgumentException("Connector's ID is null");
                }

                if (layout_connector.From == null)
                {
                    throw new System.ArgumentException("Connector's From node is null");
                }

                if (layout_connector.To == null)
                {
                    throw new System.ArgumentException("Connector's From node is null");
                }
            }
        }

        private MSAGL.Core.Layout.GeometryGraph _create_msagl_graph(DirectedGraphLayout layout_diagram)
        {
            var mg_graph = new MSAGL.Core.Layout.GeometryGraph();

            // Create the nodes in MSAGL
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var nodesize = this._to_mg_coordinates(layout_shape.Size ?? this.LayoutOptions.DefaultShapeSize);
                var node_user_data = new NodeUserData(layout_shape.ID, layout_shape);
                var center = new MSAGL.Core.Geometry.Point();
                var rectangle = MSAGL.Core.Geometry.Curves.CurveFactory.CreateRectangle(nodesize.Width, nodesize.Height, center);
                var mg_node = new MSAGL.Core.Layout.Node(rectangle, node_user_data);
                mg_graph.Nodes.Add(mg_node);
            }

            this.validate_connectors(layout_diagram);

            var mg_coordinates = this._to_mg_coordinates(this.DefaultBezierConnectorLabelBoxSize);

            var map_id_to_ud = new Dictionary<string, MSAGL.Core.Layout.Node>();
            foreach (var n in mg_graph.Nodes)
            {
                var ud = (NodeUserData)n.UserData;
                if (ud != null)
                {
                    map_id_to_ud[ud.ID] = n;
                }
            }

            // Create the MG Connectors
            foreach (var layout_connector in layout_diagram.Connectors)
            {
                if (layout_connector.From == null)
                {
                    throw new ArgumentException("Connector's From node is null");
                }

                if (layout_connector.To == null)
                {
                    throw new ArgumentException("Connector's To node is null");
                }

                var from_node = map_id_to_ud[layout_connector.From.ID];
                var to_node = map_id_to_ud[layout_connector.To.ID];

                var new_edge = new MSAGL.Core.Layout.Edge(from_node, to_node);
                // TODO: MSAGL
                //new_edge.ArrowheadAtTarget = false;
                new_edge.UserData = new NodeUserData(layout_connector.ID, layout_connector);
                mg_graph.Edges.Add(new_edge);

                new_edge.Label = new MSAGL.Core.Layout.Label(mg_coordinates.Width, mg_coordinates.Height, new_edge);
            }

            var geom_graph_components = MSAGL.Core.Layout.GraphConnectedComponents.CreateComponents(mg_graph.Nodes, mg_graph.Edges);
            var settings = new MSAGL.Layout.Layered.SugiyamaLayoutSettings();

            foreach (var subgraph in geom_graph_components)
            {
                var layout = new Microsoft.Msagl.Layout.Layered.LayeredLayout(subgraph, settings);
                subgraph.Margins = settings.NodeSeparation / 2;
                layout.Run();
            }

            // Pack the graphs using Golden Aspect Ratio
            MSAGL.Layout.MDS.MdsGraphLayout.PackGraphs(geom_graph_components, settings);

            //Update the graphs bounding box
            mg_graph.UpdateBoundingBox();

            this._mg_bb = new VA.Geometry.Rectangle(
                mg_graph.BoundingBox.Left,
                mg_graph.BoundingBox.Bottom,
                mg_graph.BoundingBox.Right,
                mg_graph.BoundingBox.Top);

            this._layout_bb = new VA.Geometry.Rectangle(0, 0, this._mg_bb.Width, this._mg_bb.Height)
                .Multiply(this.ScaleToDocument, this.ScaleToDocument);

            return mg_graph;
        }

        public void Render(
            DirectedGraphLayout layout_diagram,
            IVisio.Page page)
        {
            // Create A DOM and render it to the page
            var app = page.Application;
            var page_node = this.CreateDomPage(layout_diagram, app);

            page_node.Render(page);

            // Find all the shapes that were created in the DOM and put them in the layout structure
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var shape_node = layout_shape.DomNode;
                layout_shape.VisioShape = shape_node.VisioShape;
            }

            var layout_edges = layout_diagram.Connectors;
            foreach (var layout_edge in layout_edges)
            {
                var vnode = layout_edge.DomNode;
                layout_edge.VisioShape = vnode.VisioShape;
            }
        }

        private static U? _try_get_value<T, U>(Dictionary<T, U> dic, T t) where U : struct
        {
            U outval;
            bool r = dic.TryGetValue(t, out outval);
            if (r)
            {
                return outval;
            }
            else
            {
                return null;
            }
        }

        private static void _resolve_masters(DirectedGraphLayout layout_diagram, IVisio.Application app)
        {
            // for masters that are identified by their name and stencil, go find the actual master objects by
            // loading the specified stenciles

            var documents = app.Documents;
            var master_to_size = new Dictionary<IVisio.Master, VA.Geometry.Size>();

            // Load and cache all the masters
            var master_cache = new VA.Models.Utilities.MasterCache();
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                master_cache.Add(layout_shape.MasterName, layout_shape.StencilName);
            }
            master_cache.Resolve(documents);

            // If no size was provided for the shape, then set the size based on the master
            var layoutshapes_without_size_info = layout_diagram.Shapes.Where(s => s.Size == null);
            foreach (var layoutshape in layoutshapes_without_size_info)
            {
                var master = master_cache.Get(layoutshape.MasterName, layoutshape.StencilName);
                var size = MsaglRenderer._try_get_value(master_to_size, master.VisioMaster);
                if (!size.HasValue)
                {
                    var master_bb = master.VisioMaster.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH);
                    size = master_bb.Size;
                    master_to_size[master.VisioMaster] = size.Value;
                }
                layoutshape.Size = size.Value;
            }
        }

        public Dom.Page CreateDomPage(DirectedGraphLayout layout_diagram, IVisio.Application vis)
        {
            var page_node = new Dom.Page();
            MsaglRenderer._resolve_masters(layout_diagram, vis);

            var mg_graph = this._create_msagl_graph(layout_diagram);

            this._create_dom_shapes(page_node.Shapes, mg_graph, vis);

            if (this.LayoutOptions.UseDynamicConnectors)
            {
                this._create_dynamic_connector_edges(page_node.Shapes, mg_graph);
            }
            else
            {
                this._create_bezier_edges(page_node.Shapes, mg_graph);
            }

            // Additional Page properties
            page_node.PageLayoutCells.PlaceStyle = 1;
            page_node.PageLayoutCells.RouteStyle = 5;
            page_node.PageLayoutCells.AvenueSizeX = 2.0;
            page_node.PageLayoutCells.AvenueSizeY = 2.0;
            page_node.PageLayoutCells.LineRouteExt = 2;
            page_node.Size = this._layout_bb.Size;

            return page_node;
        }

        private void _create_dom_shapes(Dom.ShapeList domshapeslist, MSAGL.Core.Layout.GeometryGraph mg_graph, IVisio.Application app)
        {
            var node_centerpoints = mg_graph.Nodes
                    .Select(n => this._to_document_coordinates(MsaglUtil.ToVAPoint(n.Center)))
                    .ToArray();

            // Load up all the stencil docs
            var app_documents = app.Documents;
            var uds = mg_graph.Nodes.Where(n => n.UserData != null).Select(n => (NodeUserData)n.UserData).ToList();
            var shapes = uds.Where(ud => ud.Shape != null).Select(ud => ud.Shape).ToList();
            var stencilnames0 = shapes.Select(s => s.StencilName).ToList();
            var stencil_names = stencilnames0.Distinct().ToList();

            var compare = StringComparer.InvariantCultureIgnoreCase;

            var stencil_map = new Dictionary<string, IVisio.Document>(compare);
            foreach (var stencil_name in stencil_names)
            {
                if (!stencil_map.ContainsKey(stencil_name))
                {
                    var stencil = app_documents.OpenStencil(stencil_name);
                    stencil_map[stencil_name] = stencil;
                }
            }

            var master_map = new Dictionary<string, IVisio.Master>(compare);
            foreach (var nv in shapes)
            {
                var key = nv.StencilName + "+" + nv.MasterName;
                if (!master_map.ContainsKey(key))
                {
                    var stencil = stencil_map[nv.StencilName];
                    var masters = stencil.Masters;
                    var master = masters[nv.MasterName];
                    master_map[key] = master;
                }
            }

            // Create DOM Shapes for each AutoLayoutShape

            int count = 0;
            foreach (var layout_shape in shapes)
            {
                var key = layout_shape.StencilName.ToLower() + "+" + layout_shape.MasterName;
                var master = master_map[key];
                var shape_node = new Dom.Shape(master, node_centerpoints[count]);
                layout_shape.DomNode = shape_node;
                domshapeslist.Add(shape_node);
                count++;
            }

            // FORMAT EACH SHAPE
            foreach (var n in mg_graph.Nodes)
            {
                var ud = (NodeUserData)n.UserData;
                var layout_shape = ud.Shape;
                if (layout_shape != null)
                {
                    this.format_shape(layout_shape, layout_shape.DomNode);
                }
            }
        }

        private void _create_bezier_edges(Dom.ShapeList domshapes, MSAGL.Core.Layout.GeometryGraph mg_graph)
        {
            // DRAW EDGES WITH BEZIERS 
            foreach (var mg_edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)mg_edge.UserData;
                var layoutconnector = ud.Connector;
                var vconnector = this.draw_edge_bezier(mg_edge);
                layoutconnector.DomNode = vconnector;
                domshapes.Add(vconnector);
            }

            foreach (var mg_edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)mg_edge.UserData;
                var layout_connector = ud.Connector;

                if (layout_connector.Cells != null)
                {
                    var bezier_node = (Dom.BezierCurve)layout_connector.DomNode;
                    bezier_node.Cells = layout_connector.Cells.ShallowCopy();
                }
            }

            foreach (var mg_edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)mg_edge.UserData;
                var layout_connector = ud.Connector;

                if (!string.IsNullOrEmpty(layout_connector.Label))
                {
                    // this is a bezier connector
                    // draw a manual box instead
                    var label_bb = this._to_document_coordinates(MsaglUtil.ToVARectangle(mg_edge.Label.BoundingBox));
                    var vshape = new Dom.Rectangle(label_bb);
                    domshapes.Add(vshape);

                    vshape.Cells = this.DefaultBezierConnectorShapeCells.ShallowCopy();
                    vshape.Text = new VisioAutomation.Models.Text.Element(layout_connector.Label);
                }
            }
        }

        private void _create_dynamic_connector_edges(Dom.ShapeList shape_nodes, MSAGL.Core.Layout.GeometryGraph mg_graph)
        {
            // CREATE EDGES
            foreach (var edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)edge.UserData;
                var layout_connector = ud.Connector;

                Dom.Connector vconnector;
                if (layout_connector.MasterName != null && layout_connector.StencilName != null)
                {
                    vconnector = new Dom.Connector(
                    layout_connector.From.DomNode,
                    layout_connector.To.DomNode, layout_connector.MasterName, layout_connector.StencilName);
                }
                else
                {
                    
                    vconnector = new Dom.Connector(
                    layout_connector.From.DomNode,
                    layout_connector.To.DomNode, _edge_master_name, _edge_stencil_name);
                }
                layout_connector.DomNode = vconnector;
                shape_nodes.Add(vconnector);
            }

            foreach (var edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)edge.UserData;
                var layoutconnector = ud.Connector;

                var vconnector = (Dom.Connector)layoutconnector.DomNode;

                int con_route_style = (int)this.ConnectorTypeToCellVal_Appearance(layoutconnector.ConnectorType);
                int shape_route_style = (int)this.ConnectorTypeToCellVal_Style(layoutconnector.ConnectorType);

                vconnector.Text = new VisioAutomation.Models.Text.Element(layoutconnector.Label);

                vconnector.Cells = layoutconnector.Cells != null ?
                    layoutconnector.Cells.ShallowCopy()
                    : new Dom.ShapeCells();

                vconnector.Cells.ShapeLayoutConLineRouteExt = con_route_style;
                vconnector.Cells.ShapeLayoutRouteStyle = shape_route_style;
            }
        }

        private void format_shape(Shape layout_shape, Dom.BaseShape shape_node)
        {
            layout_shape.VisioShape = shape_node.VisioShape;

            // SET TEXT
            if (!string.IsNullOrEmpty(layout_shape.Label))
            {
                const char vertical_bar = '|';

                // if the shape contains vertical bars these are treated as line breaks
                if (layout_shape.Label.IndexOf(vertical_bar) >= 0)
                {
                    // there is at least one line break so this means we have to
                    // construct multiple text regions

                    // create the root text element
                    shape_node.Text = new VisioAutomation.Models.Text.Element();

                    // Split apart the string
                    var tokens = layout_shape.Label.Split(vertical_bar).Select(tok => tok.Trim()).ToArray();
                    // Add an text element for each piece
                    foreach (string token in tokens)
                    {
                        shape_node.Text.AddText(token);
                    }
                }
                else
                {
                    // No line breaks. Just use a simple TextElement with the label string
                    shape_node.Text = new VisioAutomation.Models.Text.Element(layout_shape.Label);
                }
            }

            // SET SIZE
            if (layout_shape.Size.HasValue)
            {
                shape_node.Cells.XFormWidth = layout_shape.Size.Value.Width;
                shape_node.Cells.XFormHeight = layout_shape.Size.Value.Height;
            }

            // ADD URL
            if (!string.IsNullOrEmpty(layout_shape.Url))
            {
                var hyperlink = new Dom.Hyperlink("Row_1", layout_shape.Url);
                shape_node.Hyperlinks = new List<Dom.Hyperlink> { hyperlink };
            }

            if ((layout_shape.Hyperlinks != null))
            {
                //var hyperlink = new VA.DOM.Hyperlink("Row_1", layout_shape.Url);
                shape_node.Hyperlinks = layout_shape.Hyperlinks;
            }

            // ADD CUSTOM PROPS
            if (layout_shape.CustomProperties != null)
            {
                shape_node.CustomProperties = new VA.Shapes.CustomPropertyDictionary();
                foreach (var kv in layout_shape.CustomProperties)
                {
                    shape_node.CustomProperties[kv.Key] = kv.Value;
                }
            }

            if (layout_shape.Cells != null)
            {
                shape_node.Cells = layout_shape.Cells.ShallowCopy();
            }
        }

        private Dom.BezierCurve draw_edge_bezier(MSAGL.Core.Layout.Edge edge)
        {
            var final_bez_points =
                MsaglUtil.ToVAPoints(edge).Select(p => this._to_document_coordinates(p)).ToList();

            var bez_shape = new Dom.BezierCurve(final_bez_points);
            return bez_shape;
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Appearance(ConnectorType connector_type)
        {
            switch (connector_type)
            {
                case (ConnectorType.Curved):
                    {
                        return IVisio.VisCellVals.visLORouteExtNURBS;
                    }
                case (ConnectorType.Straight):
                    {
                        return IVisio.VisCellVals.visLORouteExtStraight;
                    }
                case (ConnectorType.RightAngle):
                    {
                        return IVisio.VisCellVals.visLORouteExtStraight;
                    }
                case (ConnectorType.Default):
                    {
                        // Treat as RightAngle
                        return IVisio.VisCellVals.visLORouteExtStraight;
                    }
                default:
                    {
                        string msg = string.Format("Unhandled {0} value of {1}", nameof(ConnectorType), connector_type);
                        throw new ArgumentOutOfRangeException(nameof(connector_type), msg);
                    }
            }
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Style(ConnectorType connector_type)
        {
            switch (connector_type)
            {
                case (ConnectorType.Curved):
                    {
                        return IVisio.VisCellVals.visLORouteRightAngle;
                    }
                case (ConnectorType.Straight):
                    {
                        return IVisio.VisCellVals.visLORouteCenterToCenter;
                    }
                case (ConnectorType.RightAngle):
                    {
                        return IVisio.VisCellVals.visLORouteFlowchartNS;
                    }
                case (ConnectorType.Default):
                    {
                        return IVisio.VisCellVals.visLORouteFlowchartNS;
                    }
                default:
                    {
                        string msg = string.Format("Unhandled {0} value of {1}", nameof(ConnectorType), connector_type);
                        throw new ArgumentOutOfRangeException(nameof(connector_type), msg);
                    }
            }
        }

        public static void Render(IVisio.Page page, DirectedGraphLayout directed_graph_layout, MsaglLayoutOptions options)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            var renderer = new MsaglRenderer();
            renderer.LayoutOptions = options;
            renderer.Render(directed_graph_layout, page);
            page.ResizeToFitContents(renderer.LayoutOptions.ResizeBorderWidth);
        }
    }
}