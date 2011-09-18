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
}