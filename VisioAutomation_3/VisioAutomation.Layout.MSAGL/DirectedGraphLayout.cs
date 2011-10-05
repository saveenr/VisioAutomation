using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using MG = Microsoft.Msagl;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.MSAGL
{
    public class DirectedGraphLayout
    {
        private VA.Drawing.Rectangle msagl_bb;
        private VA.Drawing.Rectangle layout_bb;

        public VA.DOM.ShapeCells DefaultBezierConnectorShapeCells { get; set; }
        public VA.Drawing.Size DefaultBezierConnectorLabelBoxSize { get; set; }
        public VA.Layout.DirectedGraph.MSAGLLayoutOptions LayoutOptions { get; set; }

        private double ScaleToMSAGL
        {
            get { return this.LayoutOptions.ScalingFactor; }
        }

        private double ScaleToDocument
        {
            get { return 1.0/this.LayoutOptions.ScalingFactor; }
        }

        public DirectedGraphLayout()
        {
            this.LayoutOptions = new VA.Layout.DirectedGraph.MSAGLLayoutOptions();

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

        private bool validate_connectors(VA.Layout.DirectedGraph.Drawing layout_diagram)
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

        private MG.GeometryGraph CreateMSAGLGraph(VA.Layout.DirectedGraph.Drawing layout_diagram)
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

            this.msagl_bb = new VA.Drawing.Rectangle(msagl_graph.BoundingBox.Left, msagl_graph.BoundingBox.Bottom,
                                                          msagl_graph.BoundingBox.Right, msagl_graph.BoundingBox.Top);
            this.layout_bb =
                new VA.Drawing.Rectangle(0, 0, this.msagl_bb.Width, msagl_bb.Height).Multiply(
                    ScaleToDocument, ScaleToDocument);

            return msagl_graph;
        }

        // Given the MSAGL node, this function returns the Shape object
        private static VA.Layout.DirectedGraph.Shape get_shape(MG.Node msagl_node)
        {
            var shape = (VA.Layout.DirectedGraph.Shape)msagl_node.UserData;
            return shape;
        }

        internal void  _render(
            VA.Layout.DirectedGraph.Drawing layout_diagram, 
            IVisio.Page page)
        {        
            // Create A DOM and render it to the page
            var app = page.Application;
            var dom_doc = CreateDOMDocument(layout_diagram, app);
            dom_doc.ResolveVisioShapeObjects = true;

            using (var speed = new VA.FastRenderingScope(app))
            {
                dom_doc.Render(page);                    
            }

            // Find all the shapes that were created in the DOM and put them in the layout structure
            foreach (var layout_shape in layout_diagram.Shapes)
            {
                var dom_node = layout_shape.DOMNode;
                layout_shape.VisioShape = dom_node.VisioShape;
            }

            var layout_edges = layout_diagram.Connectors;
            foreach (var layout_edge in layout_edges)
            {
                var vnode = layout_edge.DOMNode;
                layout_edge.VisioShape = vnode.VisioShape;
            }
        }

        private static string handle_multiline_labels(string s)
        {
            char[] lineseps = {'|'};
            string t = s;
            t = string.Join("\n", t.Split(lineseps).Select(tok => tok.Trim()).ToArray());
            t = t.Trim();
            return t;
        }

        private static IVisio.Document OpenStencil(IVisio.Documents docs, string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            short flags = (short) IVisio.VisOpenSaveArgs.visOpenRO | (short) IVisio.VisOpenSaveArgs.visOpenDocked;
            var doc = docs.OpenEx(filename, flags);
            return doc;
        }

        private static TOut GetValueOrDefaultClass<TIn,TOut>(Dictionary<TIn,TOut> dic, TIn t) where TOut: class
        {
            TOut outval;
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

        private static TOut? GetValueOrDefaulStruct<TIn, TOut>(Dictionary<TIn, TOut> dic, TIn t) where TOut : struct 
        {
            TOut outval;
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


        private static void StoreMetadataForMasters(VA.Layout.DirectedGraph.Drawing layout_diagram, IVisio.Application app)
        {
            var documents = app.Documents;

            var name_to_stencil = new Dictionary<string, IVisio.Document>();
            var master_to_size = new Dictionary<IVisio.Master, VA.Drawing.Size>();

            // Load and cache all the masters
            var comparer = System.StringComparer.CurrentCultureIgnoreCase;
            var master_dic = new Dictionary<string, IVisio.Master>(comparer);
            var all_layout_shapes = layout_diagram.Shapes
                .Select( layout_shape => new {layout_shape, masterkey = string.Format("{0}/{1}", layout_shape.StencilName, layout_shape.MasterName)});

            // Cache all the masters based on a combinarion of their parent stencil name and the master name
            foreach (var layoutshape in all_layout_shapes)
            {
                if (!master_dic.ContainsKey(layoutshape.masterkey))
                {
                    string stencilname = layoutshape.layout_shape.StencilName.ToLower();
                    
                    var stencil = GetValueOrDefaultClass(name_to_stencil,stencilname);
                    if (stencil==null)
                    {
                        stencil = OpenStencil(documents, stencilname);
                        name_to_stencil[stencilname] = stencil;
                    }

                    IVisio.Master master = null;
                    try
                    {
                        master = stencil.Masters.ItemU[layoutshape.layout_shape.MasterName];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        string msg = string.Format("Stencil \"{0}\" does not have a Master called \"{1}\"",
                                                    stencil.Name, layoutshape.layout_shape.MasterName);
                        throw new AutomationException(msg);
                    }

                    master_dic[layoutshape.masterkey] = master;
                }
            }

            // If no size was provided for the shape, then set the size based on the master
            var layoutshapes_without_size_info = all_layout_shapes.Where(s => s.layout_shape.Size == null);
            foreach (var layoutshape in layoutshapes_without_size_info)
            {
                var master = master_dic[layoutshape.masterkey];

                var size = GetValueOrDefaulStruct(master_to_size,master);
                if (!size.HasValue)
                {
                    var master_bb = master.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH);
                    size = master_bb.Size;
                    master_to_size[master] = size.Value;
                }
                layoutshape.layout_shape.Size = size.Value;
            }
        }

        public VA.DOM.Document CreateDOMDocument(VA.Layout.DirectedGraph.Drawing layout_diagram, IVisio.Application vis)
        {
            StoreMetadataForMasters(layout_diagram, vis);

            var msagl_graph = this.CreateMSAGLGraph(layout_diagram);

            var vdoc = new VA.DOM.Document();

            vdoc.PageSettings.Size = this.layout_bb.Size;
            vdoc.PageSettings.PageCells.PlaceStyle = 1;
            vdoc.PageSettings.PageCells.RouteStyle = 5;
            vdoc.PageSettings.PageCells.AvenueSizeX = 2.0;
            vdoc.PageSettings.PageCells.AvenueSizeY = 2.0;
            vdoc.PageSettings.PageCells.LineRouteExt = 2;

            var active_window = vis.ActiveWindow;
            active_window.ShowConnectPoints = VA.Convert.BoolToShort(!this.LayoutOptions.HideConnectionPoints);
            active_window.ShowGrid = VA.Convert.BoolToShort(this.LayoutOptions.HideGrid);

            CreateDOMShapes(vdoc, msagl_graph, vis);

            if (this.LayoutOptions.UseDynamicConnectors)
            {
                CreateDynamicConnectorEdges(vdoc, msagl_graph);
            }
            else
            {
                CreateBezierEdges(vdoc, msagl_graph);
            }

            return vdoc;
        }

        private void CreateDOMShapes(VA.DOM.Document dom_doc, MG.GeometryGraph msagl_graph, IVisio.Application app)
        {
            var node_centerpoints = msagl_graph.NodeMap.Values
                    .Select(n => ToDocumentCoordinates(MSAGLUtil.ToVAPoint(n.Center)))
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
                var dom_master = new VA.DOM.Master(master, node_centerpoints[count]);
                layout_shape.DOMNode = dom_master;
                dom_doc.Shapes.Add(dom_master);
                count++;
            }

            var shape_pairs = from n in msagl_graph.NodeMap.Values
                              let ls = (VA.Layout.DirectedGraph.Shape)n.UserData
                              let vs = (VA.DOM.Shape) ls.DOMNode
                              select new {layout_shape = ls, dom_shape = vs};

            // FORMAT EACH SHAPE
            foreach (var i in shape_pairs)
            {
                format_shape(i.layout_shape, i.dom_shape);
            }
        }

        private void CreateBezierEdges(VA.DOM.Document vdoc, MG.GeometryGraph msagl_graph)
        {
// DRAW EDGES WITH BEZIERS 
            foreach (var msagl_edge in msagl_graph.Edges)
            {
                var layoutconnector = (VA.Layout.DirectedGraph.Connector)msagl_edge.UserData;
                var vconnector = draw_edge_bezier(vdoc, layoutconnector, msagl_edge);
                layoutconnector.DOMNode = vconnector;
                vdoc.Shapes.Add(vconnector);
            }

            var edge_pairs = from n in msagl_graph.Edges
                             let lc = (VA.Layout.DirectedGraph.Connector)n.UserData
                             select new { msagl_edge = n, 
                                 layout_connector = lc, 
                                 dom_bezier = (VA.DOM.BezierCurve)lc.DOMNode };

            foreach (var i in edge_pairs)
            {
                if (i.layout_connector.ShapeCells != null)
                {
                    i.dom_bezier.ShapeCells = i.layout_connector.ShapeCells.ShallowCopy();
                }
            }

            foreach (var i in edge_pairs.Where(item => !string.IsNullOrEmpty(item.layout_connector.Label)))
            {
                // this is a bezier connector
                // draw a manual box instead
                var label_bb = ToDocumentCoordinates(MSAGLUtil.ToVARectangle(i.msagl_edge.Label.BoundingBox));
                var vshape = new VA.DOM.Rectangle(label_bb);
                vdoc.Shapes.Add(vshape);

                vshape.ShapeCells = DefaultBezierConnectorShapeCells.ShallowCopy();
                vshape.Text = i.layout_connector.Label;

            }
        }

        private void CreateDynamicConnectorEdges(VA.DOM.Document vdoc, MG.GeometryGraph msagl_graph)
        {
// CREATE EDGES
            foreach (var i in msagl_graph.Edges)
            {
                var layoutconnector = (VA.Layout.DirectedGraph.Connector)i.UserData;
                var vconnector = new VA.DOM.DynamicConnector(
                    (VA.DOM.Shape)layoutconnector.From.DOMNode,
                    (VA.DOM.Shape) layoutconnector.To.DOMNode, "Dynamic Connector", "basic_u.vss");
                layoutconnector.DOMNode = vconnector;
                vdoc.Shapes.Add(vconnector);
            }

            var edge_pairs = from n in msagl_graph.Edges
                             let lc = (VA.Layout.DirectedGraph.Connector)n.UserData
                             select
                                 new { msagl_edge = n, layout_connector = lc, vconnector = (VA.DOM.DynamicConnector)lc.DOMNode };

            foreach (var i in edge_pairs)
            {
                int con_route_style = (int) dic_ct_to_appearance[i.layout_connector.ConnectorType];
                int shape_route_style = (int) dic_ct_to_style[i.layout_connector.ConnectorType];

                i.vconnector.Text = i.layout_connector.Label;

                i.vconnector.ShapeCells = i.layout_connector.ShapeCells != null ? 
                    i.layout_connector.ShapeCells.ShallowCopy()
                    : new VA.DOM.ShapeCells();

                i.vconnector.ShapeCells.ConLineRouteExt = con_route_style;
                i.vconnector.ShapeCells.ShapeRouteStyle = shape_route_style;

            }
        }

        private void format_shape(VA.Layout.DirectedGraph.Shape layout_shape, VA.DOM.Shape dom_shape)
        {
            layout_shape.VisioShape = dom_shape.VisioShape;

            // SET TEXT
            if (!string.IsNullOrEmpty(layout_shape.Label))
            {
                dom_shape.Text = handle_multiline_labels(layout_shape.Label);
            }

            // SET SIZE
            if (layout_shape.Size.HasValue)
            {
                dom_shape.ShapeCells.Width = layout_shape.Size.Value.Width;
                dom_shape.ShapeCells.Height = layout_shape.Size.Value.Height;
            }

            // ADD URL
            if (!string.IsNullOrEmpty(layout_shape.URL))
            {
                var hyperlink = new VA.DOM.Hyperlink("Row_1", layout_shape.URL);
                dom_shape.Hyperlinks = new List<VA.DOM.Hyperlink>();
                dom_shape.Hyperlinks.Add(hyperlink);
            }

            // ADD CUSTOM PROPS
            if (layout_shape.CustomProperties != null)
            {
                dom_shape.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
                foreach (var kv in layout_shape.CustomProperties)
                {
                    dom_shape.CustomProperties[kv.Key] = kv.Value;
                }
            }

            if (layout_shape.ShapeCells != null)
            {
                dom_shape.ShapeCells = layout_shape.ShapeCells.ShallowCopy();
            }
        }

        private static readonly Dictionary<VA.Connections.ConnectorType, IVisio.VisCellVals> dic_ct_to_appearance
            =
            new Dictionary<VA.Connections.ConnectorType, IVisio.VisCellVals>
                {
                    { VA.Connections.ConnectorType.Curved, IVisio.VisCellVals.visLORouteExtNURBS},
                    { VA.Connections.ConnectorType.Straight, IVisio.VisCellVals.visLORouteExtStraight},
                    { VA.Connections.ConnectorType.RightAngle, IVisio.VisCellVals.visLORouteExtDefault}
                };

        private static readonly Dictionary<VA.Connections.ConnectorType, IVisio.VisCellVals> dic_ct_to_style =
            new Dictionary<VA.Connections.ConnectorType, IVisio.VisCellVals>
                {
                    { VA.Connections.ConnectorType.Curved, IVisio.VisCellVals.visLORouteRightAngle},
                    { VA.Connections.ConnectorType.Straight, IVisio.VisCellVals.visLORouteCenterToCenter},
                    { VA.Connections.ConnectorType.RightAngle, IVisio.VisCellVals.visLORouteDefault}
                };


        private VA.DOM.BezierCurve draw_edge_bezier(
            VA.DOM.Document page,
                                            VA.Layout.DirectedGraph.Connector fc,
                                            MG.Edge edge)
        {
            var final_bez_points =
                MSAGLUtil.ToVAPoints(edge).Select(p => ToDocumentCoordinates(p)).ToList();

            var bez_shape = new VA.DOM.BezierCurve(final_bez_points);
            return bez_shape;
        }
    }
}