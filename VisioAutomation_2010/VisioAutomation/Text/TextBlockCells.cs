using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class TextBlockCells : CellGroup
    {
        public CellValueLiteral BottomMargin { get; set; }
        public CellValueLiteral LeftMargin { get; set; }
        public CellValueLiteral RightMargin { get; set; }
        public CellValueLiteral TopMargin { get; set; }
        public CellValueLiteral DefaultTabStop { get; set; }
        public CellValueLiteral Background { get; set; }
        public CellValueLiteral BackgroundTransparency { get; set; }
        public CellValueLiteral Direction { get; set; }
        public CellValueLiteral VerticalAlign { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return this.Create(nameof(this.BottomMargin), SrcConstants.TextBlockBottomMargin, this.BottomMargin);
                yield return this.Create(nameof(this.LeftMargin), SrcConstants.TextBlockLeftMargin, this.LeftMargin);
                yield return this.Create(nameof(this.RightMargin), SrcConstants.TextBlockRightMargin, this.RightMargin);
                yield return this.Create(nameof(this.TopMargin), SrcConstants.TextBlockTopMargin, this.TopMargin);
                yield return this.Create(nameof(this.DefaultTabStop), SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop);
                yield return this.Create(nameof(this.Background), SrcConstants.TextBlockBackground, this.Background);
                yield return this.Create(nameof(this.BackgroundTransparency), SrcConstants.TextBlockBackgroundTransparency, this.BackgroundTransparency);
                yield return this.Create(nameof(this.Direction), SrcConstants.TextBlockDirection, this.Direction);
                yield return this.Create(nameof(this.VerticalAlign), SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

    }
}