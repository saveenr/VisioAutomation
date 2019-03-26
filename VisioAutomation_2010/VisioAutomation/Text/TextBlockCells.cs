using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Text
{
    public class TextBlockCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral BottomMargin { get; set; }
        public VASS.CellValueLiteral LeftMargin { get; set; }
        public VASS.CellValueLiteral RightMargin { get; set; }
        public VASS.CellValueLiteral TopMargin { get; set; }
        public VASS.CellValueLiteral DefaultTabStop { get; set; }
        public VASS.CellValueLiteral Background { get; set; }
        public VASS.CellValueLiteral BackgroundTransparency { get; set; }
        public VASS.CellValueLiteral Direction { get; set; }
        public VASS.CellValueLiteral VerticalAlign { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {
                yield return this.Create(nameof(this.BottomMargin), VASS.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
                yield return this.Create(nameof(this.LeftMargin), VASS.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
                yield return this.Create(nameof(this.RightMargin), VASS.SrcConstants.TextBlockRightMargin, this.RightMargin);
                yield return this.Create(nameof(this.TopMargin), VASS.SrcConstants.TextBlockTopMargin, this.TopMargin);
                yield return this.Create(nameof(this.DefaultTabStop), VASS.SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop);
                yield return this.Create(nameof(this.Background), VASS.SrcConstants.TextBlockBackground, this.Background);
                yield return this.Create(nameof(this.BackgroundTransparency), VASS.SrcConstants.TextBlockBackgroundTransparency, this.BackgroundTransparency);
                yield return this.Create(nameof(this.Direction), VASS.SrcConstants.TextBlockDirection, this.Direction);
                yield return this.Create(nameof(this.VerticalAlign), VASS.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

    }
}