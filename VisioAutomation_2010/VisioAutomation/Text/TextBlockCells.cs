using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextBlockCells : CellRecord
    {
        public Core.CellValue BottomMargin { get; set; }
        public Core.CellValue LeftMargin { get; set; }
        public Core.CellValue RightMargin { get; set; }
        public Core.CellValue TopMargin { get; set; }
        public Core.CellValue DefaultTabStop { get; set; }
        public Core.CellValue Background { get; set; }
        public Core.CellValue BackgroundTransparency { get; set; }
        public Core.CellValue Direction { get; set; }
        public Core.CellValue VerticalAlign { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.BottomMargin), Core.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
            yield return this._create(nameof(this.LeftMargin), Core.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
            yield return this._create(nameof(this.RightMargin), Core.SrcConstants.TextBlockRightMargin, this.RightMargin);
            yield return this._create(nameof(this.TopMargin), Core.SrcConstants.TextBlockTopMargin, this.TopMargin);
            yield return this._create(nameof(this.DefaultTabStop), Core.SrcConstants.TextBlockDefaultTabStop,
                this.DefaultTabStop);
            yield return this._create(nameof(this.Background), Core.SrcConstants.TextBlockBackground, this.Background);
            yield return this._create(nameof(this.BackgroundTransparency), Core.SrcConstants.TextBlockBackgroundTransparency,
                this.BackgroundTransparency);
            yield return this._create(nameof(this.Direction), Core.SrcConstants.TextBlockDirection, this.Direction);
            yield return this._create(nameof(this.VerticalAlign), Core.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
        }


        public static CellRecords<TextBlockCells> GetTextBlockCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static TextBlockCells GetTextBlockCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static TextBlockCells RowToRecord(VisioAutomation.ShapeSheet.Data.DataRow<string> row, VisioAutomation.ShapeSheet.Data.DataColumns cols)
        {
            var record = new TextBlockCells();

            string getcellvalue(string name)
            {
                return row[cols[name].Ordinal];
            }

            record.BottomMargin = getcellvalue(nameof(TextBlockCells.BottomMargin));
            record.LeftMargin = getcellvalue(nameof(TextBlockCells.LeftMargin));
            record.RightMargin = getcellvalue(nameof(TextBlockCells.RightMargin));
            record.TopMargin = getcellvalue(nameof(TextBlockCells.TopMargin));
            record.DefaultTabStop = getcellvalue(nameof(TextBlockCells.DefaultTabStop));
            record.Background = getcellvalue(nameof(TextBlockCells.Background));
            record.BackgroundTransparency = getcellvalue(nameof(TextBlockCells.BackgroundTransparency));
            record.Direction = getcellvalue(nameof(TextBlockCells.Direction));
            record.VerticalAlign = getcellvalue(nameof(TextBlockCells.VerticalAlign));

            return record;
        }

        class Builder : CellRecordBuilder<TextBlockCells>
        {

            public Builder() : base(CellRecordQueryType.CellQuery, TextBlockCells.RowToRecord)
            {
            }
        }
    }
}