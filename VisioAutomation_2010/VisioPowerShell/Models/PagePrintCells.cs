using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PagePrintCells : VisioPowerShell.Models.BaseCells
    {
        // Page Print
        public string PrintCenterX;
        public string PrintCenterY;
        public string PrintGrid;
        public string PrintLeftMargin;
        public string PrintPageOrientation;
        public string PrintPaperKind;
        public string PrintPaperSource;
        public string PrintRightMargin;
        public string PrintScaleX;
        public string PrintScaleY;
        public string PrintTopMargin;
        public string PrintBottomMargin;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PrintCenterX), SrcConstants.PrintCenterX, this.PrintCenterX);
            yield return new CellTuple(nameof(SrcConstants.PrintCenterY), SrcConstants.PrintCenterY, this.PrintCenterY);
            yield return new CellTuple(nameof(SrcConstants.PrintGrid), SrcConstants.PrintGrid, this.PrintGrid);
            yield return new CellTuple(nameof(SrcConstants.PrintLeftMargin), SrcConstants.PrintLeftMargin, this.PrintLeftMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintPageOrientation), SrcConstants.PrintPageOrientation, this.PrintPageOrientation);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperKind), SrcConstants.PrintPaperKind, this.PrintPaperKind);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperSource), SrcConstants.PrintPaperSource, this.PrintPaperSource);
            yield return new CellTuple(nameof(SrcConstants.PrintRightMargin), SrcConstants.PrintRightMargin, this.PrintRightMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleX), SrcConstants.PrintScaleX, this.PrintScaleX);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleY), SrcConstants.PrintScaleY, this.PrintScaleY);
            yield return new CellTuple(nameof(SrcConstants.PrintTopMargin), SrcConstants.PrintTopMargin, this.PrintTopMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintBottomMargin), SrcConstants.PrintBottomMargin, this.PrintBottomMargin);
        }
    }
}