using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral PinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PinY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Angle { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormPinX, this.PinX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormPinY, this.PinY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormLocPinX, this.LocPinX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormLocPinY, this.LocPinY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormWidth, this.Width.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormHeight, this.Height.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XFormAngle, this.Angle.Value);
            }
        }

        public static List<ShapeXFormCells> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static ShapeXFormCells GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<ShapeXFormCellsReader> lazy_query = new System.Lazy<ShapeXFormCellsReader>();

        class ShapeXFormCellsReader : ReaderSingleRow<VisioAutomation.Shapes.ShapeXFormCells>
        {
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public ShapeXFormCellsReader()
            {
                this.PinX = this.query.Columns.Add(SrcConstants.XFormPinX, nameof(SrcConstants.XFormPinX));
                this.PinY = this.query.Columns.Add(SrcConstants.XFormPinY, nameof(SrcConstants.XFormPinY));
                this.LocPinX = this.query.Columns.Add(SrcConstants.XFormLocPinX, nameof(SrcConstants.XFormLocPinX));
                this.LocPinY = this.query.Columns.Add(SrcConstants.XFormLocPinY, nameof(SrcConstants.XFormLocPinY));
                this.Width = this.query.Columns.Add(SrcConstants.XFormWidth, nameof(SrcConstants.XFormWidth));
                this.Height = this.query.Columns.Add(SrcConstants.XFormHeight, nameof(SrcConstants.XFormHeight));
                this.Angle = this.query.Columns.Add(SrcConstants.XFormAngle, nameof(SrcConstants.XFormAngle));
            }

            public override ShapeXFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Shapes.ShapeXFormCells();
                cells.PinX = row[this.PinX];
                cells.PinY = row[this.PinY];
                cells.LocPinX = row[this.LocPinX];
                cells.LocPinY = row[this.LocPinY];
                cells.Width = row[this.Width];
                cells.Height = row[this.Height];
                cells.Angle = row[this.Angle];
                return cells;
            }
        }

    }
}