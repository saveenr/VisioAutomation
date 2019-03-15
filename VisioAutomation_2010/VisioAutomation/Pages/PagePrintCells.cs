using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral LeftMargin { get; set; }
        public VASS.CellValueLiteral CenterX { get; set; }
        public VASS.CellValueLiteral CenterY { get; set; }
        public VASS.CellValueLiteral OnPage { get; set; }
        public VASS.CellValueLiteral BottomMargin { get; set; }
        public VASS.CellValueLiteral RightMargin { get; set; }
        public VASS.CellValueLiteral PagesX { get; set; }
        public VASS.CellValueLiteral PagesY { get; set; }
        public VASS.CellValueLiteral TopMargin { get; set; }
        public VASS.CellValueLiteral PaperKind { get; set; }
        public VASS.CellValueLiteral Grid { get; set; }
        public VASS.CellValueLiteral Orientation { get; set; }
        public VASS.CellValueLiteral ScaleX { get; set; }
        public VASS.CellValueLiteral ScaleY { get; set; }
        public VASS.CellValueLiteral PaperSource { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LeftMargin), VASS.SrcConstants.PrintLeftMargin, this.LeftMargin);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.CenterX), VASS.SrcConstants.PrintCenterX, this.CenterX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.CenterY), VASS.SrcConstants.PrintCenterY, this.CenterY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.OnPage), VASS.SrcConstants.PrintOnPage, this.OnPage);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.BottomMargin), VASS.SrcConstants.PrintBottomMargin, this.BottomMargin);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.RightMargin), VASS.SrcConstants.PrintRightMargin, this.RightMargin);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PagesX), VASS.SrcConstants.PrintPagesX, this.PagesX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PagesY), VASS.SrcConstants.PrintPagesY, this.PagesY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.TopMargin), VASS.SrcConstants.PrintTopMargin, this.TopMargin);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PaperKind), VASS.SrcConstants.PrintPaperKind, this.PaperKind);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Grid), VASS.SrcConstants.PrintGrid, this.Grid);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Orientation), VASS.SrcConstants.PrintPageOrientation, this.Orientation);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ScaleX), VASS.SrcConstants.PrintScaleX, this.ScaleX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ScaleY), VASS.SrcConstants.PrintScaleY, this.ScaleY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PaperSource), VASS.SrcConstants.PrintPaperSource, this.PaperSource);
            }
        }

    }
}