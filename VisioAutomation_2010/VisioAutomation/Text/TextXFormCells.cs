using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellGroupBase
    {
        public CellValueLiteral Angle { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinX, this.PinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinY, this.PinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinX, this.LocPinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinY, this.LocPinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.TextXFormHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.TextXFormAngle, this.Angle);
            }
        }

        public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }

        public static TextXFormCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();


        class TextXFormCellsReader : ReaderSingleRow<Text.TextXFormCells>
        {
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public TextXFormCellsReader()
            {
                this.PinX = this.query.Columns.Add(SrcConstants.TextXFormPinX, nameof(this.PinX));
                this.PinY = this.query.Columns.Add(SrcConstants.TextXFormPinY, nameof(this.PinY));
                this.LocPinX = this.query.Columns.Add(SrcConstants.TextXFormLocPinX, nameof(this.LocPinX));
                this.LocPinY = this.query.Columns.Add(SrcConstants.TextXFormLocPinY, nameof(this.LocPinY));
                this.Width = this.query.Columns.Add(SrcConstants.TextXFormWidth, nameof(this.Width));
                this.Height = this.query.Columns.Add(SrcConstants.TextXFormHeight, nameof(this.Height));
                this.Angle = this.query.Columns.Add(SrcConstants.TextXFormAngle, nameof(this.Angle));

            }

            public override Text.TextXFormCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextXFormCells();
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