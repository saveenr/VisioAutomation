using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellGroupSingleRow
    {
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral Angle { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.XFormPinX, this.PinX);
                yield return SrcValuePair.Create(SrcConstants.XFormPinY, this.PinY);
                yield return SrcValuePair.Create(SrcConstants.XFormLocPinX, this.LocPinX);
                yield return SrcValuePair.Create(SrcConstants.XFormLocPinY, this.LocPinY);
                yield return SrcValuePair.Create(SrcConstants.XFormWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.XFormHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.XFormAngle, this.Angle);
            }
        }

        public static List<ShapeXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static ShapeXFormCells GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<ShapeXFormCellsReader> lazy_query = new System.Lazy<ShapeXFormCellsReader>();

        class ShapeXFormCellsReader : ReaderSingleRow<ShapeXFormCells>
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

            public override ShapeXFormCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new ShapeXFormCells();
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