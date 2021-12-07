using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;

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
    }
}