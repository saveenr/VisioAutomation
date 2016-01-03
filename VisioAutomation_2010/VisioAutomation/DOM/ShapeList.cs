using System.Collections.Generic;
using System.Linq;
using VACONNECT = VisioAutomation.Shapes.Connections;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Collections;

namespace VisioAutomation.DOM
{
    public class ShapeList : Node, IEnumerable<BaseShape>
    {
        private readonly NodeList<BaseShape> shapes;

        public ShapeList()
        {
            this.shapes = new NodeList<BaseShape>(this);
        }

        public IEnumerator<BaseShape> GetEnumerator()
        {
            foreach (var i in this.shapes)
            {
                yield return i;
            }
        }

        public void Add( BaseShape shape )
        {
            this.shapes.Add(shape);
        }

        public int Count
        {
            get { return this.shapes.Count; }
        }

        IEnumerator IEnumerable.GetEnumerator()    
        {                                          
            return this.GetEnumerator();
        }
        
        public void Render(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var ctx = new RenderContext(page);

            this.PrepareForDrawing(ctx);
            this.PerformDrawing(ctx);
            this.UpdateCells(ctx);
            this.SetText();
            this.SetCustomProperties(ctx);
            this.AddHyperlinks(ctx);
        }

        private void AddHyperlinks(RenderContext ctx)
        {
            var shapes_with_hyperlinks = this.shapes.Where(s => s.Hyperlinks != null);
            foreach (var shape in shapes_with_hyperlinks)
            {
                var vshape = ctx.GetShape(shape.VisioShapeID);
                foreach (var hyperlink in shape.Hyperlinks)
                {
	                    var h = vshape.Hyperlinks.Add();
                h.Name = hyperlink.Name; // Name of Hyperlink
                h.Description = hyperlink.Description;
                h.Address = hyperlink.Address; // Address of Hyperlink
                h.SubAddress = hyperlink.SubAddress;
                h.ExtraInfo = hyperlink.ExtraInfo;
                h.Frame = hyperlink.Frame;
                //h.SortKey = hyperlink.SortKey;
                //h.NewWindow = hyperlink.NewWindow;
                //h.IsDefaultLink = hyperlink.Default;
                //h.Invisible = hyperlink.Invisible;
                }
            }
        }

        private void SetCustomProperties(RenderContext ctx)
        {
            var shapes_with_custom_props = this.shapes.Where(s => s.CustomProperties != null);
            foreach (var shape in shapes_with_custom_props)
            {
                var vshape = ctx.GetShape(shape.VisioShapeID);
                foreach (var kv in shape.CustomProperties)
                {
                    string cp_name = kv.Key;
                    VACUSTPROP.CustomPropertyCells cp_cells = kv.Value;
                    VACUSTPROP.CustomPropertyHelper.Set(vshape, cp_name, cp_cells);
                }
            }
        }

        private void SetText()
        {
            var shapes_with_text = this.shapes.Where(s => s.Text != null);
            foreach (var shape in shapes_with_text)
            {
                shape.Text.SetText(shape.VisioShape);

                if (shape.TabStops != null)
                {
                    Text.TextFormat.SetTabStops(shape.VisioShape, shape.TabStops);
                }
            }
        }

        private void UpdateCells(RenderContext ctx)
        {
            this.UpdateCellsWithDropSizes(ctx);

            var update = new ShapeSheet.Update();
            var shapes_with_cells = this.shapes.Where(s => s.Cells != null);
            foreach (var shape in shapes_with_cells)
            {
                var fmt = shape.Cells;
                short id = shape.VisioShapeID;
                fmt.Apply(update, id);
            }
            update.Execute(ctx.VisioPage);
        }

        private void PerformDrawing(RenderContext ctx)
        {
            // Draw shapes
            var non_connectors = this.shapes.Where(s => !(s is Connector)).ToList();
            var non_connector_dropshapes = non_connectors.Where(s => s is Shape).Cast<Shape>().ToList();
            var non_connector_nondropshapes = non_connectors.Where(s => !(s is Shape)).ToList();

            this.drop_masters(ctx, non_connector_dropshapes);
            this._draw_non_masters(ctx, non_connector_nondropshapes);

            // verify that all non-connectors have an associated shape id
            this.check_valid_shape_ids();

            // Draw Connectors
            this._draw_connectors(ctx);

            // Make sure we have Visio shape objects for all DOM objects
            foreach (var shape in this.shapes)
            {
                if (shape.VisioShape == null)
                {
                    shape.VisioShape = ctx.GetShape(shape.VisioShapeID);
                }
            }
        }

        private void PrepareForDrawing(RenderContext ctx)
        {
            // Resolve all the masters
            this.ResolveMasters(ctx);

            // Resolve all the Character Font Name Cells
            this.ResolveFonts(ctx);
        }

        private void ResolveFonts(RenderContext ctx)
        {
            var unique_names = new HashSet<string>();
            foreach (var shape in this.shapes)
            {
                if (shape.CharFontName != null)
                {
                    if (!shape.Cells.CharFont.HasValue)
                    {
                        unique_names.Add(shape.CharFontName);
                    }
                }
            }

            var doc = ctx.VisioPage.Document;
            var fonts = doc.Fonts;

            var name_to_id = new Dictionary<string, int>(unique_names.Count);
            foreach (var name in unique_names)
            {
                // TOOD: handle exception when font is specified that does not exist
                var font = fonts[name];
                name_to_id[name] = font.ID;
            }

            foreach (var shape in this.shapes)
            {
                if (shape.CharFontName != null)
                {
                    if (!shape.Cells.CharFont.HasValue)
                    {
                        shape.Cells.CharFont = name_to_id[shape.CharFontName];
                    }
                }
            }

        }


        private void check_valid_shape_ids()
        {
            foreach (var shape in this.shapes)
            {
                if (shape is Connector)
                {
                    // do nothing
                }
                else
                {
                    if (shape.VisioShapeID < 1)
                    {
                        string msg = "A Shape drawn is missing its VisioShapeID";
                        throw new AutomationException(msg);
                    }
                }
            }
        }

        private void ResolveMasters(RenderContext ctx)
        {
            // Find all the shapes that use masters and for which
            // a Visio master object has not been identifies yet
            var shape_nodes = this.shapes
                .Where(shape => shape is Shape)
                .Cast<Shape>()
                .Where(shape => shape.Master.VisioMaster == null).ToList();

            var loader = new Internal.MasterLoader();
            foreach (var shape_node in shape_nodes)
            {
                loader.Add(shape_node.Master.MasterName,shape_node.Master.StencilName);
            }

            var application = ctx.VisioPage.Application;
            var docs = application.Documents;
            loader.Resolve(docs);

            foreach (var shape_node in shape_nodes)
            {
                var mref = loader.Get(shape_node.Master.MasterName, shape_node.Master.StencilName);
                shape_node.Master.VisioMaster = mref.VisioMaster;
            }

            // Ensure that all shapes to drop are assigned a visio master object
            foreach (var shape in this.shapes.Where(s=>s is Shape).Cast<Shape>())
            {
                if (shape.Master.VisioMaster == null)
                {
                    throw new AutomationException("Missing a master for a shape");
                }
            }
        }

        private void UpdateCellsWithDropSizes(RenderContext context)
        {
            var masters = this.shapes
                .Where(shape => shape is Shape).Cast<Shape>();

            foreach (var master in masters)
            {
                if (master.DropSize.HasValue)
                {
                    if (!master.Cells.Width.HasValue)
                    {
                        master.Cells.Width = master.DropSize.Value.Width;
                    }

                    if (!master.Cells.Height.HasValue)
                    {
                        master.Cells.Height = master.DropSize.Value.Height;
                    }
                }
            }
        }

        private void drop_masters(RenderContext ctx, List<Shape> shape_nodes)
        {
            var masters = shape_nodes.Select(m => m.Master.VisioMaster).ToList();

            var points = new List<Drawing.Point>(masters.Count);
            points.AddRange(shape_nodes.Select(s => s.DropPosition));
            var shapeids = ctx.VisioPage.DropManyU(masters, points);
            
            for (int i = 0; i < shape_nodes.Count; i++)
            {
                var master_node = shape_nodes[i];
                short shapeid = shapeids[i];
                master_node.VisioShapeID = shapeid;
            }
        }

        private void _draw_non_masters(RenderContext ctx, List<BaseShape> non_masters)
        {
            foreach (var shape in non_masters)
            {
                if (shape is Line)
                {
                    var line = (Line) shape;
                    var line_shape = ctx.VisioPage.DrawLine(line.P0, line.P1);
                    line.VisioShapeID = line_shape.ID16;
                    line.VisioShape = line_shape;
                }
                else if (shape is Rectangle)
                {
                    var rect = (Rectangle) shape;
                    var rect_shape = ctx.VisioPage.DrawRectangle(rect.P0.X, rect.P0.Y, rect.P1.X, rect.P1.Y);
                    rect.VisioShapeID = rect_shape.ID16;
                    rect.VisioShape = rect_shape;
                }
                else if (shape is Oval)
                {
                    var oval = (Oval)shape;
                    var oval_shape = ctx.VisioPage.DrawOval(oval.P0.X, oval.P0.Y, oval.P1.X, oval.P1.Y);
                    oval.VisioShapeID = oval_shape.ID16;
                    oval.VisioShape = oval_shape;
                }
                else if (shape is Arc)
                {
                    var ps = (Arc)shape;
                    var vad_arcslice = new Models.Charting.PieSlice(ps.Center, ps.StartAngle,
                                                              ps.EndAngle, ps.InnerRadius, ps.OuterRadius);
                    var ps_shape = vad_arcslice.Render(ctx.VisioPage);
                    ps.VisioShapeID = ps_shape.ID16;
                    ps.VisioShape = ps_shape;
                }
                else if (shape is PieSlice)
                {
                    var ps = (PieSlice)shape;

                    var vad_ps = new Models.Charting.PieSlice(ps.Center, ps.Start, ps.End, ps.Radius);
                    var ps_shape = vad_ps.Render(ctx.VisioPage);
                    ps.VisioShapeID = ps_shape.ID16;
                    ps.VisioShape = ps_shape;
                }
                else if (shape is BezierCurve)
                {
                    var bez = (BezierCurve) shape;
                    var bez_shape = ctx.VisioPage.DrawBezier(bez.ControlPoints);
                    bez.VisioShapeID = bez_shape.ID16;
                    bez.VisioShape = bez_shape;
                }
                else if (shape is PolyLine)
                {
                    var pl = (PolyLine) shape;
                    var pl_shape = ctx.VisioPage.DrawPolyline(pl.Points);
                    pl.VisioShapeID = pl_shape.ID16;
                    pl.VisioShape = pl_shape;
                }
                else if (shape is Connector)
                {
                    // skip these will be specially drawn
                }

                else
                {
                    string msg = $"Internal Error: Unhandled DOM node type: {shape.GetType()}";
                    throw new AutomationException(msg);
                }
            }
        }

        private void _draw_connectors(RenderContext ctx)
        {
            var connector_nodes = this.shapes.Where(s => s is Connector).Cast<Connector>().ToList();

            // if no dynamic connectors then do nothing
            if (connector_nodes.Count < 1)
            {
                return;
            }

            // Drop the number of connectors needed somewhere on the page
            var masters = connector_nodes.Select(i => i.Master.VisioMaster).ToArray();
            var origin = new Drawing.Point(-2, -2);
            var points = Enumerable.Range(0, connector_nodes.Count)
                .Select(i => origin + new Drawing.Point(1.10, 0))
                .ToList();
            var connector_shapeids = ctx.VisioPage.DropManyU(masters, points);
            var page_shapes = ctx.VisioPage.Shapes;

            // Perform the connection
            for (int i = 0; i < connector_shapeids.Length; i++)
            {
                var connector_shapeid = connector_shapeids[i];
                var vis_connector = page_shapes.ItemFromID[connector_shapeid];
                var dyncon_shape = connector_nodes[i];

                var from_shape = ctx.GetShape(dyncon_shape.From.VisioShapeID);
                var to_shape = ctx.GetShape(dyncon_shape.To.VisioShapeID);

                VACONNECT.ConnectorHelper.ConnectShapes(from_shape, to_shape, vis_connector);
                dyncon_shape.VisioShape = vis_connector;
                dyncon_shape.VisioShapeID = connector_shapeids[i];
            }
        }

        public PolyLine DrawPolyLine(IList<Drawing.Point> points)
        {
            var pl = new PolyLine(points);
            this.Add(pl);
            return pl;
        }

        public Line DrawLine(double x0, double y0, double x1, double y1)
        {
            var line = new Line(x0, y0, x1, y1);
            this.Add(line);
            return line;
        }

        public Line DrawLine(Drawing.Point p0, Drawing.Point p1)
        {
            var line = new Line(p0, p1);
            this.Add(line);
            return line;
        }

        public Rectangle DrawRectangle(double x0, double y0, double x1, double y1)
        {
            var rectangle = new Rectangle(x0, y0, x1, y1);
            this.Add(rectangle);
            return rectangle;
        }

        public Rectangle DrawRectangle(Drawing.Point p0, Drawing.Point p1)
        {
            var rectangle = new Rectangle(p0, p1);
            this.Add(rectangle);
            return rectangle;
        }


        public Oval DrawOval(Drawing.Rectangle r)
        {
            var oval = new Oval(r);
            this.Add(oval);
            return oval;
        }

        public PieSlice DrawPieSlice(Drawing.Point center, double radius, double start, double end)
        {
            var pieslice = new PieSlice(center,radius,start,end);
            this.Add(pieslice);
            return pieslice;
        }

        public Arc DrawArc(Drawing.Point center, double inner_radius, double outer_radius, double start, double end)
        {
            var arc = new Arc(center, inner_radius, outer_radius, start, end);
            this.Add(arc);
            return arc;
        }
        public Rectangle DrawRectangle(Drawing.Rectangle r)
        {
            var rectangle = new Rectangle(r);
            this.Add(rectangle);
            return rectangle;
        }

        public BezierCurve DrawBezier(IEnumerable<Drawing.Point> points)
        {
            var bezier = new BezierCurve(points);
            this.Add(bezier);
            return bezier;
        }

        public BezierCurve DrawBezier(IEnumerable<double> points)
        {
            var bezier = new BezierCurve(points);
            this.Add(bezier);
            return bezier;
        }

        public Shape Drop(IVisio.Master master, Drawing.Point pos)
        {
            var m = new Shape(master, pos);
            this.Add(m);
            return m;
        }

        public Shape Drop(IVisio.Master master, double x, double y)
        {
            var m = new Shape(master, new Drawing.Point(x, y));
            this.Add(m);
            return m;
        }

        public Shape Drop(string master, string stencil, Drawing.Point pos)
        {
            var m = new Shape(master, stencil, pos);
            this.Add(m);
            return m;
        }

        public Shape Drop(string master, string stencil, Drawing.Rectangle rect)
        {
            var m = new Shape(master, stencil, rect);
            this.Add(m);
            return m;
        }

        public Shape Drop(string master, string stencil, double x, double y)
        {
            var m = new Shape(master, stencil, new Drawing.Point(x, y));
            this.Add(m);
            return m;
        }

        public Connector Connect(IVisio.Master m, BaseShape s0, BaseShape s2)
        {
            var cxn = new Connector(s0, s2, m);
            this.Add(cxn);
            return cxn;
        }

        public Connector Connect(string master, string stencil, BaseShape s0, BaseShape s2)
        {
            var cxn = new Connector(s0, s2, master, stencil);
            this.Add(cxn);
            return cxn;
        }
    }
}