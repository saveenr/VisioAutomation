using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using MG = Microsoft.Msagl;
using VA = VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    class MsaglRenderer
    {
        private VA.Drawing.Rectangle mg_bb;
        private VA.Drawing.Rectangle layout_bb;

        public DOM.ShapeCells DefaultBezierConnectorShapeCells { get; set; }
        public VA.Drawing.Size DefaultBezierConnectorLabelBoxSize { get; set; }
        public MsaglLayoutOptions LayoutOptions { get; set; }

        private double ScaleToMsagl
        {
            get { return this.LayoutOptions.ScalingFactor; }
        }

        private double ScaleToDocument
        {
            get { return 1.0/this.LayoutOptions.ScalingFactor; }
        }

        public MsaglRenderer()
        {
            this.LayoutOptions = new MsaglLayoutOptions();

            this.DefaultBezierConnectorShapeCells = new DOM.ShapeCells();
            this.DefaultBezierConnectorShapeCells.LinePattern = 0;
            this.DefaultBezierConnectorShapeCells.LineWeight = 0.0;
            this.DefaultBezierConnectorShapeCells.FillPattern = 0;
            this.DefaultBezierConnectorLabelBoxSize = new VA.Drawing.Size(1.0, 0.5);
        }

        private VA.Drawing.Point ToDocumentCoordinates(VA.Drawing.Point point)
        {
            var np = point.Add(-this.mg_bb.Left, -this.mg_bb.Bottom).Multiply(this.ScaleToDocument);
            return np;
        }

        private VA.Drawing.Rectangle ToDocumentCoordinates(VA.Drawing.Rectangle rect)
        {
            var nr = rect.Add(-this.mg_bb.Left, -this.mg_bb.Bottom).Multiply(this.ScaleToDocument, this.ScaleToDocument);
            return nr;
        }

        private VA.Drawing.Size ToMGCoordinates(VA.Drawing.Size s)
        {
            return s.Multiply(this.ScaleToMsagl, this.ScaleToMsagl);
        }

        private void validate_connectors(Drawing layout_diagram)
        {
            foreach (var layout_connector in layout_diagram.Connectors)
            {
                if (layout_connector.ID == null)
                {
                    throw new AutomationException("Connector's ID is null");                    
                }

                if (layout_connector.From == null)
                {
                    throw new AutomationException("Connector's From node is null");
                }

                if (layout_connector.To == null)
                {
                    throw new AutomationException("Connector's From node is null");
                }
            }
        }

        private MG.Core.Layout.GeometryGraph CreateMGGraph(Drawing layout_diagram)
        {
            var mg_graph = new MG.Core.Layout.GeometryGraph();

            // Create the nodes in MSAGL
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var nodesize = this.ToMGCoordinates(layout_shape.Size ?? this.LayoutOptions.DefaultShapeSize);
                var node_user_data = new NodeUserData(layout_shape.ID, layout_shape);
                var center = new MG.Core.Geometry.Point();
                var rectangle = MG.Core.Geometry.Curves.CurveFactory.CreateRectangle(nodesize.Width, nodesize.Height, center);
                var mg_node = new MG.Core.Layout.Node( rectangle, node_user_data);
                mg_graph.Nodes.Add(mg_node);
            }

            this.validate_connectors(layout_diagram);
            
            var mg_coordinates = this.ToMGCoordinates(this.DefaultBezierConnectorLabelBoxSize);

            var map_id_to_ud = new Dictionary<string, MG.Core.Layout.Node>();
            foreach (var n in mg_graph.Nodes)
            {
                var ud = (NodeUserData) n.UserData;
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

                var new_edge = new MG.Core.Layout.Edge(from_node, to_node);
                // TODO: MSAGL
                //new_edge.ArrowheadAtTarget = false;
                new_edge.UserData = new NodeUserData(layout_connector.ID,layout_connector);
                mg_graph.Edges.Add(new_edge);

                new_edge.Label = new MG.Core.Layout.Label(mg_coordinates.Width, mg_coordinates.Height, new_edge);
            }

            var geomGraphComponents = MG.Core.Layout.GraphConnectedComponents.CreateComponents(mg_graph.Nodes, mg_graph.Edges);
            var settings = new MG.Layout.Layered.SugiyamaLayoutSettings();

            foreach (var subgraph in geomGraphComponents)
            {
                var layout = new Microsoft.Msagl.Layout.Layered.LayeredLayout(subgraph, settings);
                subgraph.Margins = settings.NodeSeparation / 2;
                layout.Run();
            }

            // Pack the graphs using Golden Aspect Ratio
            MG.Layout.MDS.MdsGraphLayout.PackGraphs(geomGraphComponents, settings);

            //Update the graphs bounding box
            mg_graph.UpdateBoundingBox();

            this.mg_bb = new VA.Drawing.Rectangle(
                mg_graph.BoundingBox.Left, 
                mg_graph.BoundingBox.Bottom,
                mg_graph.BoundingBox.Right,
                mg_graph.BoundingBox.Top);
            
            this.layout_bb = new VA.Drawing.Rectangle(0, 0, this.mg_bb.Width, this.mg_bb.Height)
                .Multiply(this.ScaleToDocument, this.ScaleToDocument);

            return mg_graph;
        }

        public void  Render(
            Drawing layout_diagram, 
            IVisio.Page page)
        {        
            // Create A DOM and render it to the page
            var app = page.Application;
            var page_node = this.CreateDOMPage(layout_diagram, app);

            page_node.Render(page);                    

            // Find all the shapes that were created in the DOM and put them in the layout structure
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var shape_node = layout_shape.DOMNode;
                layout_shape.VisioShape = shape_node.VisioShape;
            }

            var layout_edges = layout_diagram.Connectors;
            foreach (var layout_edge in layout_edges)
            {
                var vnode = layout_edge.DOMNode;
                layout_edge.VisioShape = vnode.VisioShape;
            }
        }

        private static U? TryGetValue<T, U>(Dictionary<T, U> dic, T t) where U : struct 
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

        private static void ResolveMasters(Drawing layout_diagram, IVisio.Application app)
        {
            // for masters that are identified by their name and stencil, go find the actual master objects by
            // loading the specified stenciles

            var documents = app.Documents;
            var master_to_size = new Dictionary<IVisio.Master, VA.Drawing.Size>();

            // Load and cache all the masters
            var loader = new Internal.MasterLoader();
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                loader.Add(layout_shape.MasterName,layout_shape.StencilName);                
            }
            loader.Resolve(documents);
            
            // If no size was provided for the shape, then set the size based on the master
            var layoutshapes_without_size_info = layout_diagram.Shapes.Where(s => s.Size == null);
            foreach (var layoutshape in layoutshapes_without_size_info)
            {
                var master = loader.Get(layoutshape.MasterName,layoutshape.StencilName);
                var size = MsaglRenderer.TryGetValue(master_to_size,master.VisioMaster);
                if (!size.HasValue)
                {
                    var master_bb = master.VisioMaster.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH);
                    size = master_bb.Size;
                    master_to_size[master.VisioMaster] = size.Value;
                }
                layoutshape.Size = size.Value;
            }
        }

        public DOM.Page CreateDOMPage(Drawing layout_diagram, IVisio.Application vis)
        {
            var page_node = new DOM.Page();
            MsaglRenderer.ResolveMasters(layout_diagram, vis);

            var mg_graph = this.CreateMGGraph(layout_diagram);

            this.CreateDOMShapes(page_node.Shapes, mg_graph, vis);



            if (this.LayoutOptions.UseDynamicConnectors)
            {
                this.CreateDynamicConnectorEdges(page_node.Shapes, mg_graph);
            }
            else
            {
                this.CreateBezierEdges(page_node.Shapes, mg_graph);
            }


            // Additional Page properties
            page_node.PageCells.PlaceStyle = 1;
            page_node.PageCells.RouteStyle = 5;
            page_node.PageCells.AvenueSizeX = 2.0;
            page_node.PageCells.AvenueSizeY = 2.0;
            page_node.PageCells.LineRouteExt = 2;
            page_node.Size = this.layout_bb.Size;

            return page_node;
        }

        private void CreateDOMShapes(DOM.ShapeList domshapeslist, MG.Core.Layout.GeometryGraph mg_graph, IVisio.Application app)
        {
            var node_centerpoints = mg_graph.Nodes
                    .Select(n => this.ToDocumentCoordinates(Internal.MsaglUtil.ToVAPoint(n.Center)))
                    .ToArray();

            // Load up all the stencil docs
            var app_documents = app.Documents;
            var uds = mg_graph.Nodes.Where(n => n.UserData != null).Select(n => (NodeUserData) n.UserData).ToList();
            var shapes = uds.Where(ud => ud.Shape != null).Select(ud=>ud.Shape).ToList();
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
                var shape_node = new DOM.Shape(master, node_centerpoints[count]);
                layout_shape.DOMNode = shape_node;
                domshapeslist.Add(shape_node);
                count++;
            }

            // FORMAT EACH SHAPE
            foreach (var n in mg_graph.Nodes)
            {
                var ud = (NodeUserData) n.UserData;
                var layout_shape = ud.Shape;
                if (layout_shape != null)
                {
                    this.format_shape(layout_shape, layout_shape.DOMNode);                    
                }
            }
        }

        private void CreateBezierEdges(DOM.ShapeList domshapes, MG.Core.Layout.GeometryGraph mg_graph)
        {
            // DRAW EDGES WITH BEZIERS 
            foreach (var mg_edge in mg_graph.Edges)
            {
                var ud = (NodeUserData) mg_edge.UserData;
                var layoutconnector =  ud.Connector;
                var vconnector = this.draw_edge_bezier(mg_edge);
                layoutconnector.DOMNode = vconnector;
                domshapes.Add(vconnector);
            }

            foreach (var mg_edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)mg_edge.UserData;
                var layout_connector = ud.Connector;

                if (layout_connector.Cells != null)
                {
                    var bezier_node = (DOM.BezierCurve)layout_connector.DOMNode;
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
                    var label_bb = this.ToDocumentCoordinates(Internal.MsaglUtil.ToVARectangle(mg_edge.Label.BoundingBox));
                    var vshape = new DOM.Rectangle(label_bb);
                    domshapes.Add(vshape);

                    vshape.Cells = this.DefaultBezierConnectorShapeCells.ShallowCopy();
                    vshape.Text = new Text.Markup.TextElement(layout_connector.Label);                    
                }
            }
        }

        private void CreateDynamicConnectorEdges(DOM.ShapeList shape_nodes, MG.Core.Layout.GeometryGraph mg_graph)
        {
            // CREATE EDGES
            foreach (var edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)edge.UserData;
                var layout_connector = ud.Connector;

              VisioAutomation.DOM.Connector vconnector;
              if (layout_connector.MasterName != null && layout_connector.StencilName != null)
              {
                  vconnector = new VA.DOM.Connector(
                  layout_connector.From.DOMNode,
                  layout_connector.To.DOMNode, layout_connector.MasterName, layout_connector.StencilName);
              }
              else
              {
                  vconnector = new VA.DOM.Connector(
                  layout_connector.From.DOMNode,
                  layout_connector.To.DOMNode, "Dynamic Connector", "connec_u.vss");
              }
                layout_connector.DOMNode = vconnector;
                shape_nodes.Add(vconnector);
            }

            foreach (var edge in mg_graph.Edges)
            {
                var ud = (NodeUserData)edge.UserData;
                var layoutconnector = ud.Connector;

                var vconnector = (DOM.Connector) layoutconnector.DOMNode;

                int con_route_style = (int) this.ConnectorTypeToCellVal_Appearance(layoutconnector.ConnectorType);
                int shape_route_style = (int) this.ConnectorTypeToCellVal_Style(layoutconnector.ConnectorType);

                vconnector.Text = new Text.Markup.TextElement(layoutconnector.Label);

                vconnector.Cells = layoutconnector.Cells != null ? 
                    layoutconnector.Cells.ShallowCopy()
                    : new DOM.ShapeCells();

                vconnector.Cells.ConLineRouteExt = con_route_style;
                vconnector.Cells.ShapeRouteStyle = shape_route_style;
            }
        }

        private void format_shape(Shape layout_shape, DOM.BaseShape shape_node)
        {
            layout_shape.VisioShape = shape_node.VisioShape;

            // SET TEXT
            if (!string.IsNullOrEmpty(layout_shape.Label))
            {
                // if the shape contains vertical bars these are treated as line breaks
                if (layout_shape.Label.IndexOf('|') >= 0)
                {
                    // there is at least one line break so this means we have to
                    // construct multiple text regions

                    // create the root text element
                    shape_node.Text = new Text.Markup.TextElement();

                    // Split apart the string
                    var tokens = layout_shape.Label.Split('|').Select(tok => tok.Trim()).ToArray();
                    // Add an text element for each piece
                    foreach (string token in tokens)
                    {
                        shape_node.Text.AddText(token);
                    }
                }
                else
                {
                    // No line braeaks. Just use a simple TextElement with the label string
                    shape_node.Text = new Text.Markup.TextElement(layout_shape.Label);
                }
            }

            // SET SIZE
            if (layout_shape.Size.HasValue)
            {
                shape_node.Cells.Width = layout_shape.Size.Value.Width;
                shape_node.Cells.Height = layout_shape.Size.Value.Height;
            }

            // ADD URL
            if (!string.IsNullOrEmpty(layout_shape.URL))
            {
                var hyperlink = new DOM.Hyperlink("Row_1", layout_shape.URL);
                shape_node.Hyperlinks = new List<DOM.Hyperlink> {hyperlink};
            }

            if ((layout_shape.Hyperlinks != null))
            {
                //var hyperlink = new VA.DOM.Hyperlink("Row_1", layout_shape.URL);
                shape_node.Hyperlinks = layout_shape.Hyperlinks;
            }

            // ADD CUSTOM PROPS
            if (layout_shape.CustomProperties != null)
            {
                shape_node.CustomProperties = new Dictionary<string, Shapes.CustomProperties.CustomPropertyCells>();
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

        private DOM.BezierCurve draw_edge_bezier(MG.Core.Layout.Edge edge)
        {
            var final_bez_points =
                Internal.MsaglUtil.ToVAPoints(edge).Select(p => this.ToDocumentCoordinates(p)).ToList();

            var bez_shape = new DOM.BezierCurve(final_bez_points);
            return bez_shape;
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Appearance(Shapes.Connections.ConnectorType ct)
        {
            switch (ct)
            {
                case (Shapes.Connections.ConnectorType.Curved):
                    {
                        return IVisio.VisCellVals.visLORouteExtNURBS;
                    }
                case (Shapes.Connections.ConnectorType.Straight):
                    {
                        return IVisio.VisCellVals.visLORouteExtStraight;
                    }
                case (Shapes.Connections.ConnectorType.RightAngle):
                    {
                        return IVisio.VisCellVals.visLORouteExtStraight;
                    }
                default:
                    {
                        throw new ArgumentOutOfRangeException();
                    }
            }
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Style(Shapes.Connections.ConnectorType ct)
        {
            switch (ct)
            {
                case (Shapes.Connections.ConnectorType.Curved):
                    {
                        return IVisio.VisCellVals.visLORouteRightAngle;
                    }
                case (Shapes.Connections.ConnectorType.Straight):
                    {
                        return IVisio.VisCellVals.visLORouteCenterToCenter;
                    }
                case (Shapes.Connections.ConnectorType.RightAngle):
                    {
                        return IVisio.VisCellVals.visLORouteFlowchartNS;
                    }
                default:
                    {
                        throw new ArgumentOutOfRangeException();
                    }
            }
        }

        public static void Render(IVisio.Page page, Drawing drawing, MsaglLayoutOptions options)
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
            renderer.Render(drawing,page);
            page.ResizeToFitContents(renderer.LayoutOptions.ResizeBorderWidth);
        }
    }
}