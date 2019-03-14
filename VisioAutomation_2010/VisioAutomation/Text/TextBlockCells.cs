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

        public override IEnumerable<NamedSrcValuePair> NamedSrcValuePairs
        {
            get
            {


                yield return NamedSrcValuePair.Create(nameof(this.BottomMargin), SrcConstants.TextBlockBottomMargin, this.BottomMargin);
                yield return NamedSrcValuePair.Create(nameof(this.LeftMargin), SrcConstants.TextBlockLeftMargin, this.LeftMargin);
                yield return NamedSrcValuePair.Create(nameof(this.RightMargin), SrcConstants.TextBlockRightMargin, this.RightMargin);
                yield return NamedSrcValuePair.Create(nameof(this.TopMargin), SrcConstants.TextBlockTopMargin, this.TopMargin);
                yield return NamedSrcValuePair.Create(nameof(this.DefaultTabStop), SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop);
                yield return NamedSrcValuePair.Create(nameof(this.Background), SrcConstants.TextBlockBackground, this.Background);
                yield return NamedSrcValuePair.Create(nameof(this.BackgroundTransparency), SrcConstants.TextBlockBackgroundTransparency, this.BackgroundTransparency);
                yield return NamedSrcValuePair.Create(nameof(this.Direction), SrcConstants.TextBlockDirection, this.Direction);
                yield return NamedSrcValuePair.Create(nameof(this.VerticalAlign), SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

    }
}