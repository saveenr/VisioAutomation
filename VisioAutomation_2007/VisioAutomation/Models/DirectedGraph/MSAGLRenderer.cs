using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes.Connections;
using VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using MG = Microsoft.Msagl;
using VA = VisioAutomation;
using DGMODEL = VisioAutomation.Models.DirectedGraph;

namespace VisioAutomation.Models.DirectedGraph
{
    class MSAGLRenderer
    {
        private VA.Drawing.Rectangle msagl_bb;
        private VA.Drawing.Rectangle layout_bb;

        public VA.DOM.ShapeCells DefaultBezierConnectorShapeCells { get; set; }
        public VA.Drawing.Size DefaultBezierConnectorLabelBoxSize { get; set; }
        public DGMODEL.MSAGLLayoutOptions LayoutOptions { get; set; }

        private double ScaleToMSAGL
        {
            get { return this.LayoutOptions.ScalingFactor; }
        }

        private double ScaleToDocument
        {
            get { return 1.0/this.LayoutOptions.ScalingFactor; }
        }

        public MSAGLRenderer()
        {
            this.LayoutOptions = new DGMODEL.MSAGLLayoutOptions();

            this.DefaultBezierConnectorShapeCells = new VA.DOM.ShapeCells();
            DefaultBezierConnectorShapeCells.LinePattern = 0;
            DefaultBezierConnectorShapeCells.LineWeight = 0.0;
            DefaultBezierConnectorShapeCells.FillPattern = 0;
            DefaultBezierConnectorLabelBoxSize = new VA.Drawing.Size(1.0, 0.5);
        }

        private VA.Drawing.Point ToDocumentCoordinates(VA.Drawing.Point point)
        {
            var np = point.Add(-msagl_bb.Left, -msagl_bb.Bottom).Multiply(ScaleToDocument);
            return np;
        }

        private VA.Drawing.Rectangle ToDocumentCoordinates(VA.Drawing.Rectangle rect)
        {
            var nr = rect.Add(-msagl_bb.Left, -msagl_bb.Bottom).Multiply(ScaleToDocument,
                                                                      ScaleToDocument);
            return nr;
        }

        private VA.Drawing.Size ToMSAGLCoordinates(VA.Drawing.Size s)
        {
            return s.Multiply(ScaleToMSAGL, ScaleToMSAGL);
        }

        private bool validate_connectors(DGMODEL.Drawing layout_diagram)
        {
            bool success = true;
            foreach (var layout_connector in layout_diagram.Connectors)
            {
                if (layout_connector.From == null)
                {
                    throw new VA.AutomationException("Connector's From node is null");
                }

                if (layout_connector.To == null)
                {
                    throw new VA.AutomationException("Connector's From node is null");
                }
            }

            return success;
        }

        private MG.GeometryGraph CreateMSAGLGraph(DGMODEL.Drawing layout_diagram)
        {
            var msagl_graph = new MG.GeometryGraph();
            var defsize = new VA.Drawing.Size(this.LayoutOptions.DefaultShapeSize.Width,
                                                   this.LayoutOptions.DefaultShapeSize.Height);

            // Create the nodes in MSAGL
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var nodesize = ToMSAGLCoordinates(layout_shape.Size ?? defsize);
                var msagl_node = new MG.Node(layout_shape.ID,
                                             MG.Splines.CurveFactory.CreateBox(nodesize.Width, nodesize.Height,
                                                                               new MG.Point()));
                msagl_graph.AddNode(msagl_node);
                msagl_node.UserData = layout_shape;
            }

            bool connectors_ok = this.validate_connectors(layout_diagram);
            // TODO: What to do if connectors_ok is false?

            var msagl_size = this.ToMSAGLCoordinates(DefaultBezierConnectorLabelBoxSize);

            // Create the MSAGL Connectors
            foreach (var layout_connector in layout_diagram.Connectors)
            {
                if (layout_connector.From == null)
                {
                    throw new System.ArgumentException("Connector's From node is null");
                }

                if (layout_connector.To == null)
                {
                    throw new System.ArgumentException("Connector's To node is null");
                }

                var from_node = msagl_graph.NodeMap[layout_connector.From.ID];
                var to_node = msagl_graph.NodeMap[layout_connector.To.ID];

                var new_edge = new MG.Edge(from_node, to_node);
                new_edge.ArrowheadAtTarget = false;
                new_edge.UserData = layout_connector;
                msagl_graph.AddEdge(new_edge);

                new_edge.Label = new Microsoft.Msagl.Label(msagl_size.Width, msagl_size.Height, new_edge);
            }

            msagl_graph.CalculateLayout();

            this.msagl_bb = new VA.Drawing.Rectangle(
                msagl_graph.BoundingBox.Left, 
                msagl_graph.BoundingBox.Bottom,
                msagl_graph.BoundingBox.Right,
                msagl_graph.BoundingBox.Top);
            
            this.layout_bb = new VA.Drawing.Rectangle(0, 0, this.msagl_bb.Width, msagl_bb.Height)
                .Multiply(ScaleToDocument, ScaleToDocument);

            return msagl_graph;
        }

        // Given the MSAGL node, this function returns the Shape object
        private static DGMODEL.Shape get_shape(MG.Node msagl_node)
        {
            var shape = (DGMODEL.Shape)msagl_node.UserData;
            return shape;
        }

        public void  Render(
            DGMODEL.Drawing layout_diagram, 
            IVisio.Page page)
        {        
            // Create A DOM and render it to the page
            var app = page.Application;
            var page_node = CreateDOMPage(layout_diagram, app);

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

        private static void ResolveMasters(DGMODEL.Drawing layout_diagram, IVisio.Application app)
        {
            // for masters that are identified by their name and stencil, go find the actual master objects by
            // loading the specified stenciles

            var documents = app.Documents;
            var master_to_size = new Dictionary<IVisio.Master, VA.Drawing.Size>();

            // Load and cache all the masters
            var loader = new VA.Internal.MasterLoader();
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
                var size = TryGetValue(master_to_size,master.VisioMaster);
                if (!size.HasValue)
                {
                    var master_bb = master.VisioMaster.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH);
                    size = master_bb.Size;
                    master_to_size[master.VisioMaster] = size.Value;
                }
                layoutshape.Size = size.Value;
            }
        }

        public VA.DOM.Page CreateDOMPage(DGMODEL.Drawing layout_diagram, IVisio.Application vis)
        {
            var page_node = new VA.DOM.Page();
            ResolveMasters(layout_diagram, vis);

            var msagl_graph = this.CreateMSAGLGraph(layout_diagram);

            CreateDOMShapes(page_node.Shapes, msagl_graph, vis);

            if (this.LayoutOptions.UseDynamicConnectors)
            {
                CreateDynamicConnectorEdges(page_node.Shapes, msagl_graph);
            }
            else
            {
                CreateBezierEdges(page_node.Shapes, msagl_graph);
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

        private void CreateDOMShapes(VA.DOM.ShapeList domshapeslist, MG.GeometryGraph msagl_graph, IVisio.Application app)
        {
            var node_centerpoints = msagl_graph.NodeMap.Values
                    .Select(n => ToDocumentCoordinates(VA.Internal.MSAGLUtil.ToVAPoint(n.Center)))
                    .ToArray();

            // Load up all the stencil docs
            var app_documents = app.Documents;
            var nodes = msagl_graph.NodeMap.Values.Select(get_shape);
            var stencil_names = nodes.Select(s => s.StencilName.ToUpper()).Distinct().ToList();
            
            var stencil_map = new Dictionary<string,IVisio.Document>();
            foreach (var stencil_name in stencil_names)
            {
                if (!stencil_map.ContainsKey(stencil_name))
                {
                    var stencil = app_documents.OpenStencil(stencil_name);
                    stencil_map[stencil_name] = stencil;
                }
            }

            var master_map = new Dictionary<string,IVisio.Master>();
            foreach (var nv in nodes)
            {
                var key = nv.StencilName.ToLower() + "+" + nv.MasterName; 
                if (!master_map.ContainsKey(key))
                {
                    var stencil = stencil_map[nv.StencilName.ToUpper()];
                    var masters = stencil.Masters;
                    var master = masters[nv.MasterName];
                    master_map[key] = master;
                }
            }

            // Create DOM Shapes for each AutoLayoutShape

            int count = 0;
            foreach (var layout_shape in nodes)
            {
                var key = layout_shape.StencilName.ToLower() + "+" + layout_shape.MasterName;
                var master = master_map[key];
                var shape_node = new VA.DOM.Shape(master, node_centerpoints[count]);
                layout_shape.DOMNode = shape_node;
                domshapeslist.Add(shape_node);
                count++;
            }

            var shape_pairs = from n in msagl_graph.NodeMap.Values
                              let layout_shape = (DGMODEL.Shape)n.UserData
                              select new
                                  {
                                      layout_shape,
                                      shape_node = (VA.DOM.BaseShape)layout_shape.DOMNode
                                  };

            // FORMAT EACH SHAPE
            foreach (var i in shape_pairs)
            {
                format_shape(i.layout_shape, i.shape_node);
            }
        }

        private void CreateBezierEdges(VA.DOM.ShapeList domshapes, MG.GeometryGraph msagl_graph)
        {
            // DRAW EDGES WITH BEZIERS 
            foreach (var msagl_edge in msagl_graph.Edges)
            {
                var layoutconnector = (DGMODEL.Connector)msagl_edge.UserData;
                var vconnector = draw_edge_bezier(domshapes, layoutconnector, msagl_edge);
                layoutconnector.DOMNode = vconnector;
                domshapes.Add(vconnector);
            }

            var edge_pairs = from n in msagl_graph.Edges
                             let lc = (DGMODEL.Connector)n.UserData
                             select new { msagl_edge = n, 
                                 layout_connector = lc, 
                                 bezier_node = (VA.DOM.BezierCurve)lc.DOMNode };

            foreach (var i in edge_pairs)
            {
                if (i.layout_connector.Cells != null)
                {
                    i.bezier_node.Cells = i.layout_connector.Cells.ShallowCopy();
                }
            }

            foreach (var i in edge_pairs.Where(item => !string.IsNullOrEmpty(item.layout_connector.Label)))
            {
                // this is a bezier connector
                // draw a manual box instead
                var label_bb = ToDocumentCoordinates(VA.Internal.MSAGLUtil.ToVARectangle(i.msagl_edge.Label.BoundingBox));
                var vshape = new VA.DOM.Rectangle(label_bb);
                domshapes.Add(vshape);

                vshape.Cells = DefaultBezierConnectorShapeCells.ShallowCopy();
                vshape.Text = new VA.Text.Markup.TextElement(i.layout_connector.Label);

            }
        }

        private void CreateDynamicConnectorEdges(VA.DOM.ShapeList shape_nodes, MG.GeometryGraph msagl_graph)
        {
            // CREATE EDGES
            foreach (var i in msagl_graph.Edges)
            {
                var layoutconnector = (DGMODEL.Connector)i.UserData;
                var vconnector = new VA.DOM.Connector(
                    layoutconnector.From.DOMNode,
                    layoutconnector.To.DOMNode, "Dynamic Connector", "connec_u.vss");
                layoutconnector.DOMNode = vconnector;
                shape_nodes.Add(vconnector);
            }

            var edge_pairs = from n in msagl_graph.Edges
                             let lc = (DGMODEL.Connector)n.UserData
                             select
                                 new { msagl_edge = n, layout_connector = lc, vconnector = (VA.DOM.Connector)lc.DOMNode };

            foreach (var i in edge_pairs)
            {
                int con_route_style = (int)  ConnectorTypeToCellVal_Appearance(i.layout_connector.ConnectorType);
                int shape_route_style = (int) ConnectorTypeToCellVal_Style(i.layout_connector.ConnectorType);

                i.vconnector.Text = new VA.Text.Markup.TextElement(i.layout_connector.Label);

                i.vconnector.Cells = i.layout_connector.Cells != null ? 
                    i.layout_connector.Cells.ShallowCopy()
                    : new VA.DOM.ShapeCells();

                i.vconnector.Cells.ConLineRouteExt = con_route_style;
                i.vconnector.Cells.ShapeRouteStyle = shape_route_style;

            }
        }

        private void format_shape(DGMODEL.Shape layout_shape, VA.DOM.BaseShape shape_node)
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
                    shape_node.Text = new VA.Text.Markup.TextElement();

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
                    shape_node.Text = new VA.Text.Markup.TextElement(layout_shape.Label);
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
                var hyperlink = new VA.DOM.Hyperlink("Row_1", layout_shape.URL);
                shape_node.Hyperlinks = new List<VA.DOM.Hyperlink> {hyperlink};
            }

            // ADD CUSTOM PROPS
            if (layout_shape.CustomProperties != null)
            {
                shape_node.CustomProperties = new Dictionary<string, CustomPropertyCells>();
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

        private VA.DOM.BezierCurve draw_edge_bezier(
            VA.DOM.ShapeList page,
            DGMODEL.Connector connector,
            MG.Edge edge)
        {
            var final_bez_points =
                VA.Internal.MSAGLUtil.ToVAPoints(edge).Select(p => ToDocumentCoordinates(p)).ToList();

            var bez_shape = new VA.DOM.BezierCurve(final_bez_points);
            return bez_shape;
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Appearance(ConnectorType ct)
        {
            switch (ct)
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
                        return IVisio.VisCellVals.visLORouteExtDefault;
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
            }
        }

        private IVisio.VisCellVals ConnectorTypeToCellVal_Style(ConnectorType ct)
        {
            switch (ct)
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
                        return IVisio.VisCellVals.visLORouteDefault;
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
            }
        }

        public static void Render(IVisio.Page page, VisioAutomation.Models.DirectedGraph.Drawing drawing, DGMODEL.MSAGLLayoutOptions options)
        {
            var renderer = new VA.Models.DirectedGraph.MSAGLRenderer();
            renderer.LayoutOptions = options;
            renderer.Render(drawing,page);
            page.ResizeToFitContents(renderer.LayoutOptions.ResizeBorderWidth);
        }
    }
}