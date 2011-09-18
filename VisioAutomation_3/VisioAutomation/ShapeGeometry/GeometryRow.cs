using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public abstract class GeometryRow
    {
        public abstract void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section);

    }

    public class MoveToRow : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }

        internal MoveToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }

        public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagMoveTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, this.X);
            update.SetFormulaIgnoreNull(y_src, this.Y);
        }
    }

    public class LineToRow : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }

        internal LineToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }

        public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagLineTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, this.X);
            update.SetFormulaIgnoreNull(y_src, this.Y);
        }


    }

    public class ArcToRow : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
        public VA.ShapeSheet.FormulaLiteral A { get; private set; }

        internal ArcToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            this.X = x;
            this.Y = y;
            this.A = a;
        }

        public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagArcTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, this.X);
            update.SetFormulaIgnoreNull(y_src, this.Y);
            update.SetFormulaIgnoreNull(a_src, this.A);
        }
    }

    public class EllipticalArcToRow : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
        public VA.ShapeSheet.FormulaLiteral A { get; private set; }
        public VA.ShapeSheet.FormulaLiteral B { get; private set; }
        public VA.ShapeSheet.FormulaLiteral C { get; private set; }
        public VA.ShapeSheet.FormulaLiteral D { get; private set; }

        internal EllipticalArcToRow(
             VA.ShapeSheet.FormulaLiteral x, 
             VA.ShapeSheet.FormulaLiteral y, 
             VA.ShapeSheet.FormulaLiteral a,
             VA.ShapeSheet.FormulaLiteral b,
             VA.ShapeSheet.FormulaLiteral c,
             VA.ShapeSheet.FormulaLiteral d)
        {
            this.X = x;
            this.Y = y;
            this.A = a;
            this.B = b;
            this.C = c;
            this.D = d;
        }

        public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagArcTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
            var b_src = VA.ShapeSheet.SRCConstants.Geometry_B.ForSectionAndRow(section, row_index);
            var c_src = VA.ShapeSheet.SRCConstants.Geometry_C.ForSectionAndRow(section, row_index);
            var d_src = VA.ShapeSheet.SRCConstants.Geometry_D.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, this.X);
            update.SetFormulaIgnoreNull(y_src, this.Y);
            update.SetFormulaIgnoreNull(a_src, this.A);
            update.SetFormulaIgnoreNull(b_src, this.B);
            update.SetFormulaIgnoreNull(c_src, this.C);
            update.SetFormulaIgnoreNull(d_src, this.D);
        }
    }

    public class EllipseRow : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
        public VA.ShapeSheet.FormulaLiteral A { get; private set; }
        public VA.ShapeSheet.FormulaLiteral B { get; private set; }
        public VA.ShapeSheet.FormulaLiteral C { get; private set; }
        public VA.ShapeSheet.FormulaLiteral D { get; private set; }

        internal EllipseRow(
             VA.ShapeSheet.FormulaLiteral x,
             VA.ShapeSheet.FormulaLiteral y,
             VA.ShapeSheet.FormulaLiteral a,
             VA.ShapeSheet.FormulaLiteral b,
             VA.ShapeSheet.FormulaLiteral c,
             VA.ShapeSheet.FormulaLiteral d)
        {
            this.X = x;
            this.Y = y;
            this.A = a;
            this.B = b;
            this.C = c;
            this.D = d;
        }

        public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagArcTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
            var b_src = VA.ShapeSheet.SRCConstants.Geometry_B.ForSectionAndRow(section, row_index);
            var c_src = VA.ShapeSheet.SRCConstants.Geometry_C.ForSectionAndRow(section, row_index);
            var d_src = VA.ShapeSheet.SRCConstants.Geometry_D.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, this.X);
            update.SetFormulaIgnoreNull(y_src, this.Y);
            update.SetFormulaIgnoreNull(a_src, this.A);
            update.SetFormulaIgnoreNull(b_src, this.B);
            update.SetFormulaIgnoreNull(c_src, this.C);
            update.SetFormulaIgnoreNull(d_src, this.D);
        }

        public class NURBSToRow : VA.ShapeGeometry.GeometryRow
        {
            public VA.ShapeSheet.FormulaLiteral X { get; private set; }
            public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
            public VA.ShapeSheet.FormulaLiteral A { get; private set; }
            public VA.ShapeSheet.FormulaLiteral B { get; private set; }
            public VA.ShapeSheet.FormulaLiteral C { get; private set; }
            public VA.ShapeSheet.FormulaLiteral D { get; private set; }
            public VA.ShapeSheet.FormulaLiteral E { get; private set; }

            internal NURBSToRow(
                 VA.ShapeSheet.FormulaLiteral x,
                 VA.ShapeSheet.FormulaLiteral y,
                 VA.ShapeSheet.FormulaLiteral a,
                 VA.ShapeSheet.FormulaLiteral b,
                 VA.ShapeSheet.FormulaLiteral c,
                 VA.ShapeSheet.FormulaLiteral d,
                 VA.ShapeSheet.FormulaLiteral e)
            {
                this.X = x;
                this.Y = y;
                this.A = a;
                this.B = b;
                this.C = c;
                this.D = d;
            }

            public override void AddToShape(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
            {
                short row_index = shape.AddRow(section, row, (short)IVisio.VisRowTags.visTagArcTo);
                var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
                var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
                var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
                var b_src = VA.ShapeSheet.SRCConstants.Geometry_B.ForSectionAndRow(section, row_index);
                var c_src = VA.ShapeSheet.SRCConstants.Geometry_C.ForSectionAndRow(section, row_index);
                var d_src = VA.ShapeSheet.SRCConstants.Geometry_D.ForSectionAndRow(section, row_index);
                var e_src = VA.ShapeSheet.SRCConstants.Geometry_E.ForSectionAndRow(section, row_index);
                update.SetFormulaIgnoreNull(x_src, this.X);
                update.SetFormulaIgnoreNull(y_src, this.Y);
                update.SetFormulaIgnoreNull(a_src, this.A);
                update.SetFormulaIgnoreNull(b_src, this.B);
                update.SetFormulaIgnoreNull(c_src, this.C);
                update.SetFormulaIgnoreNull(d_src, this.D);
                update.SetFormulaIgnoreNull(e_src, this.E);
            }
        }
    }

}