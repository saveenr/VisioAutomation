using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Text
{
    public class TextBlockCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.BottomMargin), Core.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
            yield return this.Create(nameof(this.LeftMargin), Core.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
            yield return this.Create(nameof(this.RightMargin), Core.SrcConstants.TextBlockRightMargin, this.RightMargin);
            yield return this.Create(nameof(this.TopMargin), Core.SrcConstants.TextBlockTopMargin, this.TopMargin);
            yield return this.Create(nameof(this.DefaultTabStop), Core.SrcConstants.TextBlockDefaultTabStop,
                this.DefaultTabStop);
            yield return this.Create(nameof(this.Background), Core.SrcConstants.TextBlockBackground, this.Background);
            yield return this.Create(nameof(this.BackgroundTransparency), Core.SrcConstants.TextBlockBackgroundTransparency,
                this.BackgroundTransparency);
            yield return this.Create(nameof(this.Direction), Core.SrcConstants.TextBlockDirection, this.Direction);
            yield return this.Create(nameof(this.VerticalAlign), Core.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
        }
    }
}