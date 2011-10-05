using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.DOM
{
    public class Document : Node
    {
        public NodeList<Shape> Shapes { get; private set; }
        public PageSettings PageSettings { get; set; }
        public bool ResolveVisioShapeObjects { get; set; }

        public Document()
        {
            this.Shapes = new NodeList<Shape>(this);
            this.PageSettings = new PageSettings(8.5, 11);
        }

        public override IEnumerable<Node> Children
        {
            get { return Shapes.Cast<Node>(); }
        }

        public void Render(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            this._Render(page);
        }

        public void _Render(IVisio.Page page)
        {
            // ----------------------------------------
            // Preparation

            var ctx = new RenderContext(page);

            // Resolve all the masters

            LoadMastersDeferred(ctx);

            // ----------------------------------------
            // Handle the initial page settings
            // Set the page properties before the rest of the shapes are dropped
            initialize_page(ctx);

            // ----------------------------------------
            // Draw shapes

            foreach (var cat_shapes in VA.Internal.LinqUtil.ChunkByBool(this.Shapes, s=>s is Master))
            {
                var masters_col = new List<Master>();
                var shapes_col = new List<Shape>();
                if (cat_shapes.Items.Count > 0)
                {
                    if (cat_shapes.Category == true)
                    {
                        // true means this is a master
                        masters_col.Clear();
                        masters_col.AddRange( cat_shapes.Items.Cast<Master>());
                        _draw_masters(ctx,masters_col);
                        masters_col.Clear();
                    }
                    else
                    {
                        shapes_col.Clear();
                        shapes_col.AddRange( cat_shapes.Items);
                        _draw_non_masters(ctx,shapes_col);
                        shapes_col.Clear();
                    }
                }
            }


            // verify that all non-connectors have an associated shape id
            check_valid_shape_ids();

            // Draw Connectors
            _draw_dynamic_connectors(ctx);

            // If needed get all shape objects
            if (this.ResolveVisioShapeObjects)
            {
                foreach (var shape in this.Shapes)
                {
                    if (shape.VisioShape == null)
                    {
                        shape.VisioShape = ctx.GetShapeObjectForID(shape.VisioShapeID);
                    }
                }
            }

            // ----------------------------------------
            // Set Shape format on all shapes
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            var shapes_with_formatting = this.Shapes.Where(s => s.ShapeCells != null);
            foreach (var shape in shapes_with_formatting)
            {
                var fmt = shape.ShapeCells;
                short id = shape.VisioShapeID;
                fmt.Apply(update, id);
            }
            update.Execute(page);

            // ----------------------------------------
            // set the shape text
            var shapes_with_text = this.Shapes.Where(s => s.Text != null);
            foreach (var shape in shapes_with_text)
            {
                var vshape = ctx.GetShapeObjectForID(shape.VisioShapeID);
                vshape.Text = shape.Text;

                if (shape.TextElement != null)
                {
                    shape.TextElement.SetShapeText(shape.VisioShape);
                }
                else if (shape.Text != null)
                {
                    vshape.Text = shape.Text;
                }
                else
                {
                    // do nothing there is no text
                }

                if (shape.TabStops != null)
                {
                    VA.Text.TextHelper.SetTabStops(shape.VisioShape, shape.TabStops);
                }
            }

            // ----------------------------------------
            // Apply Custom Properties
            var shapes_with_custom_props = this.Shapes.Where(s => s.CustomProperties != null);
            foreach (var shape in shapes_with_custom_props)
            {
                var vshape = ctx.GetShapeObjectForID(shape.VisioShapeID);
                foreach (var kv in shape.CustomProperties)
                {
                    string cp_name = kv.Key;
                    VA.CustomProperties.CustomPropertyCells cp_cells = kv.Value;
                    VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(vshape, cp_name, cp_cells);
                }
            }

            // ----------------------------------------
            // Apply Hyperlinks Properties
            var shapes_with_hyperlinks = this.Shapes.Where(s => s.Hyperlinks != null);
            foreach (var shape in shapes_with_hyperlinks)
            {
                var vshape = ctx.GetShapeObjectForID(shape.VisioShapeID);
                foreach (var hyperlink in shape.Hyperlinks)
                {
                    var h = vshape.Hyperlinks.Add();
                    h.Name = hyperlink.Name; // Name of Hyperlink
                    h.Address = hyperlink.Address; // Address of Hyperlink
                }
            }
        }

        private void initialize_page(RenderContext ctx)
        {
            ctx.VisioPage.Name = this.PageSettings.Name;
            if (this.PageSettings.Size.HasValue)
            {
                ctx.VisioPage.SetSize(this.PageSettings.Size.Value);
            }
            var page_sheet = ctx.VisioPage.PageSheet;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            this.PageSettings.PageCells.Apply(update, (short) page_sheet.ID);
            update.Execute(ctx.VisioPage);
        }

        private void check_valid_shape_ids()
        {
            foreach (var shape in this.Shapes)
            {
                if (shape is DynamicConnector)
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

        internal class MasterLoader
        {
            public class MasterRef
            {
                public string StencilName;
                public string MasterName;
                public IVisio.Master VisioMaster;
            }

            private Dictionary<string, MasterRef> dic;

            public MasterLoader()
            {
                this.dic = new Dictionary<string, MasterRef>();
            }

            public void Add(string mastername, string stencilname)
            {
                var item = new MasterRef();
                item.MasterName = mastername;
                item.StencilName = stencilname;
                item.VisioMaster = null;

                string key = this.getkey(mastername, stencilname);
                this.dic[key] = item;
            }

            private string getkey(string mastername, string stencilname)
            {
                return mastername + "+" + stencilname;
            }

            public MasterRef Get(string mastername, string stencilname)
            {
                string key = this.getkey(mastername, stencilname);
                return this.dic[key];
            }

            public bool Contains(string mastername, string stencilname)
            {
                string key = this.getkey(mastername, stencilname);
                return this.dic.ContainsKey(key);
            }


            public void Resolve(IVisio.Documents docs)
            {
                var unique_stencils = new HashSet<string>();
                foreach (var kv in this.dic)
                {
                    string mr = kv.Value.StencilName;
                    unique_stencils.Add(mr);
                }
                var name_to_stencildoc = new Dictionary<string, IVisio.Document>();
                foreach (var stencil in unique_stencils)
                {
                    try
                    {
                        var stencil_doc = docs.OpenStencil(stencil);
                        name_to_stencildoc[stencil] = stencil_doc;
                    }
                    catch (Exception)
                    {
                        string msg = string.Format("No such Stencil \"{0}\"", stencil);
                        throw new AutomationException(msg);
                    }
                }

                // identify real master objects for all deferred shapes
                foreach (var mr in this.dic.Values)
                {
                    if (mr.VisioMaster == null)
                    {
                        var stencildoc = name_to_stencildoc[mr.StencilName];
                        var stencilmasters = stencildoc.Masters;

                        try
                        {
                            var master_object = stencilmasters[mr.MasterName];
                            mr.VisioMaster= master_object;
                        }
                        catch (Exception)
                        {
                            string msg = string.Format("No such Master \"{0}\" in Stencil \"{1}\"",
                                                       mr.MasterName, mr.StencilName);
                            throw new AutomationException(msg);
                        }
                    }
                }               
            }

        }

        private void LoadMastersDeferred(RenderContext ctx)
        {
            var deferred_shapes = this.Shapes
                .Where(shape => shape is ShapeFromMaster)
                .Cast<ShapeFromMaster>()
                .Where(shape => shape.MasterObject == null);

            var mres = new MasterLoader();
            foreach (var s in deferred_shapes)
            {
                mres.Add(s.MasterName,s.StencilName);
            }

            var application = ctx.VisioPage.Application;
            var docs = application.Documents;
            mres.Resolve(docs);

            foreach (var s in deferred_shapes)
            {
                var mref = mres.Get(s.MasterName, s.StencilName);
                s.MasterObject = mref.VisioMaster;
            }

            // Ensure that all masters have objects now
            foreach (var deferred_shape in deferred_shapes)
            {
                if (deferred_shape.MasterObject == null)
                {
                    throw new AutomationException("Found master without stencil object");
                }
            }
        }

        private void _draw_masters(RenderContext ctx, List<Master> dom_masters)
        {
            var masters = dom_masters.Select(m => m.MasterObject).ToList();

            var points = new List<VA.Drawing.Point>(masters.Count);
            points.AddRange(dom_masters.Select(s => s.DropPosition));
            var shapeids = ctx.VisioPage.DropManyU(masters, points);
            
            for (int i = 0; i < dom_masters.Count; i++)
            {
                var dom_master = dom_masters[i];
                short shapeid = shapeids[i];
                dom_master.VisioShapeID = shapeid;
            }
        }

        private void _draw_non_masters(RenderContext ctx, List<Shape> non_masters)
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
                    var vad_arcslice = new VA.Layout.Radial.DoughnutSlice(ps.Center, ps.StartAngle,
                                                              ps.EndAngle, ps.InnerRadius, ps.OuterRadius);
                    var ps_shape = vad_arcslice.Render(ctx.VisioPage);
                    ps.VisioShapeID = ps_shape.ID16;
                    ps.VisioShape = ps_shape;
                }
                else if (shape is PieSlice)
                {
                    var ps = (PieSlice)shape;

                    var vad_ps = new VA.Layout.Radial.PieSlice(ps.Center, ps.Start, ps.End, ps.Radius);
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
                else if (shape is DynamicConnector)
                {
                    // skip these will be specially drawn
                }

                else
                {
                    string msg = string.Format("Unhandled Node Type: {0}", shape.GetType());
                    throw new AutomationException(msg);
                }
            }
        }

        private void _draw_dynamic_connectors(RenderContext ctx)
        {
            var dyncon_shapes = this.Shapes.Where(s => s is DynamicConnector).Cast<DynamicConnector>().ToList();
            if (dyncon_shapes.Count > 0)
            {
                var masterobjects = dyncon_shapes.Select(i => i.MasterObject).ToArray();
                var origin = new VA.Drawing.Point(-2, -2);
                var points = Enumerable.Range(0, dyncon_shapes.Count)
                    .Select(i => origin + new VA.Drawing.Point(1.10, 0))
                    .ToList();

                var shapeids = ctx.VisioPage.DropManyU(masterobjects, points);

                for (int i = 0; i < shapeids.Length; i++)
                {
                    var connector_id = shapeids[i];
                    var page_shapes = ctx.VisioPage.Shapes;
                    var vis_connector = page_shapes.ItemFromID[connector_id];
                    var dyncon_shape = dyncon_shapes[i];

                    var from_shape = ctx.GetShapeObjectForID(dyncon_shape.From.VisioShapeID);
                    var to_shape = ctx.GetShapeObjectForID(dyncon_shape.To.VisioShapeID);
                    VA.Connections.ConnectorHelper.ConnectShapes(vis_connector, from_shape, to_shape);
                    dyncon_shape.VisioShape = vis_connector;
                    dyncon_shape.VisioShapeID = shapeids[i];
                }
            }
        }

        public Line DrawLine(double x0, double y0, double x1, double y1)
        {
            var line = new Line(x0, y0, x1, y1);
            this.Shapes.Add(line);
            return line;
        }

        public Line DrawLine(VA.Drawing.Point p0, VA.Drawing.Point p1)
        {
            var line = new Line(p0, p1);
            this.Shapes.Add(line);
            return line;
        }

        public Rectangle DrawRectangle(double x0, double y0, double x1, double y1)
        {
            var rectangle = new Rectangle(x0, y0, x1, y1);
            this.Shapes.Add(rectangle);
            return rectangle;
        }

        public Rectangle DrawRectangle(VA.Drawing.Point p0, VA.Drawing.Point p1)
        {
            var rectangle = new Rectangle(p0, p1);
            this.Shapes.Add(rectangle);
            return rectangle;
        }


        public Oval DrawOval(VA.Drawing.Rectangle r)
        {
            var oval = new Oval(r);
            this.Shapes.Add(oval);
            return oval;
        }

        public PieSlice DrawPieSlice(VA.Drawing.Point center, double radius, double start, double end)
        {
            var pieslice = new PieSlice(center,radius,start,end);
            this.Shapes.Add(pieslice);
            return pieslice;
        }

        public Arc DrawArc(VA.Drawing.Point center, double inner_radius, double outer_radius, double start, double end)
        {
            var arc = new Arc(center, inner_radius, outer_radius, start, end);
            this.Shapes.Add(arc);
            return arc;
        }
        public Rectangle DrawRectangle(VA.Drawing.Rectangle r)
        {
            var rectangle = new Rectangle(r);
            this.Shapes.Add(rectangle);
            return rectangle;
        }

        public BezierCurve DrawBezier(IEnumerable<VA.Drawing.Point> points)
        {
            var bezier = new BezierCurve(points);
            this.Shapes.Add(bezier);
            return bezier;
        }

        public BezierCurve DrawBezier(IEnumerable<double> points)
        {
            var bezier = new BezierCurve(points);
            this.Shapes.Add(bezier);
            return bezier;
        }

        public Master Drop(IVisio.Master master, VA.Drawing.Point pos)
        {
            var m = new Master(master, pos);
            this.Shapes.Add(m);
            return m;
        }

        public Master Drop(IVisio.Master master, double x, double y)
        {
            var m = new Master(master, new VA.Drawing.Point(x, y));
            this.Shapes.Add(m);
            return m;
        }

        public Master Drop(string master, string stencil, VA.Drawing.Point pos)
        {
            var m = new Master(master, stencil, pos);
            this.Shapes.Add(m);
            return m;
        }

        public Master Drop(string master, string stencil, double x, double y)
        {
            var m = new Master(master, stencil, new VA.Drawing.Point(x, y));
            this.Shapes.Add(m);
            return m;
        }

        public DynamicConnector Connect(IVisio.Master m, Shape s0, Shape s2)
        {
            var dc = new DynamicConnector(s0, s2, m);
            this.Shapes.Add(dc);
            return dc;
        }

        public DynamicConnector Connect(string master, string stencil, Shape s0, Shape s2)
        {
            var dc = new DynamicConnector(s0, s2, master, stencil);
            this.Shapes.Add(dc);
            return dc;
        }
    }
}