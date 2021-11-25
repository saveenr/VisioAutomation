using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Text
{
    public class TextBlockCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue BottomMargin { get; set; }
        public VisioAutomation.Core.CellValue LeftMargin { get; set; }
        public VisioAutomation.Core.CellValue RightMargin { get; set; }
        public VisioAutomation.Core.CellValue TopMargin { get; set; }
        public VisioAutomation.Core.CellValue DefaultTabStop { get; set; }
        public VisioAutomation.Core.CellValue Background { get; set; }
        public VisioAutomation.Core.CellValue BackgroundTransparency { get; set; }
        public VisioAutomation.Core.CellValue Direction { get; set; }
        public VisioAutomation.Core.CellValue VerticalAlign { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.BottomMargin), VisioAutomation.Core.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
            yield return this.Create(nameof(this.LeftMargin), VisioAutomation.Core.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
            yield return this.Create(nameof(this.RightMargin), VisioAutomation.Core.SrcConstants.TextBlockRightMargin, this.RightMargin);
            yield return this.Create(nameof(this.TopMargin), VisioAutomation.Core.SrcConstants.TextBlockTopMargin, this.TopMargin);
            yield return this.Create(nameof(this.DefaultTabStop), VisioAutomation.Core.SrcConstants.TextBlockDefaultTabStop,
                this.DefaultTabStop);
            yield return this.Create(nameof(this.Background), VisioAutomation.Core.SrcConstants.TextBlockBackground, this.Background);
            yield return this.Create(nameof(this.BackgroundTransparency), VisioAutomation.Core.SrcConstants.TextBlockBackgroundTransparency,
                this.BackgroundTransparency);
            yield return this.Create(nameof(this.Direction), VisioAutomation.Core.SrcConstants.TextBlockDirection, this.Direction);
            yield return this.Create(nameof(this.VerticalAlign), VisioAutomation.Core.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
        }
    }
}