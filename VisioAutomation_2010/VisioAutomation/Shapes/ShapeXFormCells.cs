using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellGroup
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

        public static List<ShapeXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeXFormCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeXFormCellsReader> lazy_reader = new System.Lazy<ShapeXFormCellsReader>();

        class ShapeXFormCellsReader : CellGroupReader<ShapeXFormCells>
        {
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public ShapeXFormCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.PinX = this.query_singlerow.Columns.Add(SrcConstants.XFormPinX, nameof(this.PinX));
                this.PinY = this.query_singlerow.Columns.Add(SrcConstants.XFormPinY, nameof(this.PinY));
                this.LocPinX = this.query_singlerow.Columns.Add(SrcConstants.XFormLocPinX, nameof(this.LocPinX));
                this.LocPinY = this.query_singlerow.Columns.Add(SrcConstants.XFormLocPinY, nameof(this.LocPinY));
                this.Width = this.query_singlerow.Columns.Add(SrcConstants.XFormWidth, nameof(this.Width));
                this.Height = this.query_singlerow.Columns.Add(SrcConstants.XFormHeight, nameof(this.Height));
                this.Angle = this.query_singlerow.Columns.Add(SrcConstants.XFormAngle, nameof(this.Angle));
            }

            public override ShapeXFormCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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