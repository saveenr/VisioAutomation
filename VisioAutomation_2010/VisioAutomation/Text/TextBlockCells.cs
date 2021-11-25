using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text;

public class TextBlockCells : VASS.CellGroups.CellGroup
{
    public VASS.CellValue BottomMargin { get; set; }
    public VASS.CellValue LeftMargin { get; set; }
    public VASS.CellValue RightMargin { get; set; }
    public VASS.CellValue TopMargin { get; set; }
    public VASS.CellValue DefaultTabStop { get; set; }
    public VASS.CellValue Background { get; set; }
    public VASS.CellValue BackgroundTransparency { get; set; }
    public VASS.CellValue Direction { get; set; }
    public VASS.CellValue VerticalAlign { get; set; }

    public override IEnumerable<CellMetadataItem> GetCellMetadata()
    {
        yield return this.Create(nameof(this.BottomMargin), VASS.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
        yield return this.Create(nameof(this.LeftMargin), VASS.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
        yield return this.Create(nameof(this.RightMargin), VASS.SrcConstants.TextBlockRightMargin, this.RightMargin);
        yield return this.Create(nameof(this.TopMargin), VASS.SrcConstants.TextBlockTopMargin, this.TopMargin);
        yield return this.Create(nameof(this.DefaultTabStop), VASS.SrcConstants.TextBlockDefaultTabStop,
            this.DefaultTabStop);
        yield return this.Create(nameof(this.Background), VASS.SrcConstants.TextBlockBackground, this.Background);
        yield return this.Create(nameof(this.BackgroundTransparency), VASS.SrcConstants.TextBlockBackgroundTransparency,
            this.BackgroundTransparency);
        yield return this.Create(nameof(this.Direction), VASS.SrcConstants.TextBlockDirection, this.Direction);
        yield return this.Create(nameof(this.VerticalAlign), VASS.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
    }
}