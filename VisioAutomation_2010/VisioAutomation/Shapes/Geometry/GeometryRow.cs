using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Shapes.Geometry
{
    public class GeometryRow
    {
        public ShapeSheet.ValueLiteral X { get; set; }
        public ShapeSheet.ValueLiteral Y { get; set; }
        public ShapeSheet.ValueLiteral A { get; set; }
        public ShapeSheet.ValueLiteral B { get; set; }
        public ShapeSheet.ValueLiteral C { get; set; }
        public ShapeSheet.ValueLiteral D { get; set; }
        public ShapeSheet.ValueLiteral E { get; set; }
        public IVisio.VisRowTags RowTag { get; set; }

        public GeometryRow(IVisio.VisRowTags tag)
        {
            this.RowTag = tag;
        }

        public virtual IVisio.VisRowTags GetRowTagType()
        {
            return this.RowTag;
        }

        public void AddTo(IVisio.Shape shape, ShapeSheet.Writer.ShapeSheetWriter writer, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) this.GetRowTagType());
            this.Update(section, row_index, writer);
        }


        private void Update(short section_index, short row_index, ShapeSheet.Writer.ShapeSheetWriter writer)
        {
            var x_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_X.Cell);
            var y_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_Y.Cell);
            var a_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_A.Cell);
            var b_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_B.Cell);
            var c_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_C.Cell);
            var d_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_D.Cell);
            var e_src = new VA.ShapeSheet.SRC(section_index, row_index,ShapeSheet.SRCConstants.Geometry_E.Cell);

            writer.SetFormula(x_src, this.X);
            writer.SetFormula(y_src, this.Y);
            writer.SetFormula(a_src, this.A);
            writer.SetFormula(b_src, this.B);
            writer.SetFormula(c_src, this.C);
            writer.SetFormula(d_src, this.D);
            writer.SetFormula(e_src, this.E);
        }

        public static GeometryRow CreateLineTo(ShapeSheet.ValueLiteral x, ShapeSheet.ValueLiteral y)
        {
            // http://msdn.microsoft.com/en-us/library/aa195656(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagLineTo);
            row.X = x;
            row.Y = y;
            return row;
        }

        public static GeometryRow CreateMoveTo(ShapeSheet.ValueLiteral x, ShapeSheet.ValueLiteral y)
        {
            // http://msdn.microsoft.com/en-us/library/aa195679(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagMoveTo);
            row.X = x;
            row.Y = y;
            return row;
        }

        public static GeometryRow CreateArcTo(ShapeSheet.ValueLiteral x, ShapeSheet.ValueLiteral y,
                                              ShapeSheet.ValueLiteral a)
        {
            // http://msdn.microsoft.com/en-us/library/aa195698(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagArcTo);
            row.X = x;
            row.Y = y;
            row.A = a;
            return row;
        }

        public static GeometryRow CreateEllipticalArcTo(ShapeSheet.ValueLiteral x,
                                                        ShapeSheet.ValueLiteral y,
                                                        ShapeSheet.ValueLiteral a,
                                                        ShapeSheet.ValueLiteral b,
                                                        ShapeSheet.ValueLiteral c,
                                                        ShapeSheet.ValueLiteral d)
        {
            // http://msdn.microsoft.com/en-us/library/aa195660(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagEllipticalArcTo);
            row.X = x;
            row.Y = y;
            row.A = a;
            row.B = b;
            row.C = c;
            row.D = d;
            return row;
        }

        public static GeometryRow CreateEllipse(ShapeSheet.ValueLiteral x,
                                                ShapeSheet.ValueLiteral y,
                                                ShapeSheet.ValueLiteral a,
                                                ShapeSheet.ValueLiteral b,
                                                ShapeSheet.ValueLiteral c,
                                                ShapeSheet.ValueLiteral d)
        {

            // http://msdn.microsoft.com/en-us/library/aa195692(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagEllipse);
            row.X = x;
            row.Y = y;
            row.A = a;
            row.B = b;
            row.C = c;
            row.D = d;
            return row;
        }

        public static GeometryRow CreateNURBSTo(ShapeSheet.ValueLiteral x,
                                                ShapeSheet.ValueLiteral y,
                                                ShapeSheet.ValueLiteral a,
                                                ShapeSheet.ValueLiteral b,
                                                ShapeSheet.ValueLiteral c,
                                                ShapeSheet.ValueLiteral d,
                                                ShapeSheet.ValueLiteral e)
        {
            // http://msdn.microsoft.com/en-us/library/aa195685(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagEllipse);
            row.X = x;
            row.Y = y;
            row.A = a;
            row.B = b;
            row.C = c;
            row.D = d;
            row.E = e;
            return row;
        }

        public static GeometryRow CreatePolylineTo(ShapeSheet.ValueLiteral x,
                                        ShapeSheet.ValueLiteral y,
                                        ShapeSheet.ValueLiteral a)
        {
            // http://msdn.microsoft.com/en-us/library/aa195682(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagPolylineTo);
            row.X = x;
            row.Y = y;
            row.A = a;
            return row;
        }

        public static GeometryRow CreateInfiniteLine(ShapeSheet.ValueLiteral x,
                                ShapeSheet.ValueLiteral y,
                                ShapeSheet.ValueLiteral a,
                                ShapeSheet.ValueLiteral b)
        {
            // http://msdn.microsoft.com/en-us/library/aa195682(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagInfiniteLine);
            row.X = x;
            row.Y = y;
            row.A = a;
            row.B = b;
            return row;
        }

        public static GeometryRow CreateSplineKnot(ShapeSheet.ValueLiteral x,
                                ShapeSheet.ValueLiteral y,
                                ShapeSheet.ValueLiteral a)
        {
            // http://msdn.microsoft.com/en-us/library/aa195667(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagSplineSpan);
            row.X = x;
            row.Y = y;
            row.A = a;
            return row;
        }

        public static GeometryRow CreateSplineStart(ShapeSheet.ValueLiteral x,
                                                ShapeSheet.ValueLiteral y,
                                                ShapeSheet.ValueLiteral a,
                                                ShapeSheet.ValueLiteral b,
                                                ShapeSheet.ValueLiteral c,
                                                ShapeSheet.ValueLiteral d)
        {

            // http://msdn.microsoft.com/en-us/library/aa195663(v=office.11).aspx

            var row = new GeometryRow(IVisio.VisRowTags.visTagSplineBeg);
            row.X = x;
            row.Y = y;
            row.A = a;
            row.B = b;
            row.C = c;
            row.D = d;
            return row;
        }
    }
}