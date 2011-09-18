using VA=VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public class GeometryRow
    {
        
    }

    public class RowMoveTo : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }

        internal RowMoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }
    }

    public class RowLineTo : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }

        internal RowLineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            this.X = x;
            this.Y = y;
        }
    }

    public class RowArcTo : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
        public VA.ShapeSheet.FormulaLiteral A { get; private set; }

        internal RowArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            this.X = x;
            this.Y = y;
            this.A = a;
        }
    }

    public class RowEllipticalArcTo : VA.ShapeGeometry.GeometryRow
    {
        public VA.ShapeSheet.FormulaLiteral X { get; private set; }
        public VA.ShapeSheet.FormulaLiteral Y { get; private set; }
        public VA.ShapeSheet.FormulaLiteral A { get; private set; }
        public VA.ShapeSheet.FormulaLiteral B { get; private set; }
        public VA.ShapeSheet.FormulaLiteral C { get; private set; }
        public VA.ShapeSheet.FormulaLiteral D { get; private set; }

        internal RowEllipticalArcTo(
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
    }
}