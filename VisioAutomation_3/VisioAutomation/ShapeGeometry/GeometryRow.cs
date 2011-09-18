using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public abstract class GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral A { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral B { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral C { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral D { get; protected set; }
        public VA.ShapeSheet.FormulaLiteral E { get; protected set; }

        public abstract IVisio.VisRowTags GetRowTagType();

        public void AddTo(IVisio.Shape shape, VA.ShapeSheet.Update.SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) this.GetRowTagType());
            this.Update(section,row_index,update);
        }

        private void Update(short section, short row_index, VA.ShapeSheet.Update.SRCUpdate update)
        {
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

    public class MoveToRow : VA.ShapeGeometry.GeometryRow
    {
        internal MoveToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }

        public override IVisio.VisRowTags  GetRowTagType()
        {
            return IVisio.VisRowTags.visTagMoveTo;
        }
    }

    public class LineToRow : VA.ShapeGeometry.GeometryRow
    {
        internal LineToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }

        public override IVisio.VisRowTags GetRowTagType()
        {
            return IVisio.VisRowTags.visTagLineTo;
        }
    }

    public class ArcToRow : VA.ShapeGeometry.GeometryRow
    {
        internal ArcToRow(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            this.X = x;
            this.Y = y;
            this.A = a;
        }

        public override IVisio.VisRowTags GetRowTagType()
        {
            return IVisio.VisRowTags.visTagArcTo;
        }
    }

    public class EllipticalArcToRow : VA.ShapeGeometry.GeometryRow
    {
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

        public override IVisio.VisRowTags GetRowTagType()
        {
            return IVisio.VisRowTags.visTagEllipticalArcTo;
        }
    }

    public class EllipseRow : VA.ShapeGeometry.GeometryRow
    {

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

        public override IVisio.VisRowTags GetRowTagType()
        {
            return IVisio.VisRowTags.visTagEllipse;
        }

        public class NURBSToRow : VA.ShapeGeometry.GeometryRow
        {
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

            public override IVisio.VisRowTags GetRowTagType()
            {
                return IVisio.VisRowTags.visTagNURBSTo;
            }
        }
    }

}