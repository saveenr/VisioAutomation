using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PagePrintCells : VisioPowerShell.Models.BaseCells
    {
        public string CenterX;
        public string CenterY;
        public string PrintGrid;
        public string LeftMargin;
        public string PageOrientation;
        public string PaperKind;
        public string PaperSource;
        public string RightMargin;
        public string ScaleX;
        public string ScaleY;
        public string TopMargin;
        public string BottomMargin;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PrintCenterX), SrcConstants.PrintCenterX, this.CenterX);
            yield return new CellTuple(nameof(SrcConstants.PrintCenterY), SrcConstants.PrintCenterY, this.CenterY);
            yield return new CellTuple(nameof(SrcConstants.PrintGrid), SrcConstants.PrintGrid, this.PrintGrid);
            yield return new CellTuple(nameof(SrcConstants.PrintLeftMargin), SrcConstants.PrintLeftMargin, this.LeftMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintPageOrientation), SrcConstants.PrintPageOrientation, this.PageOrientation);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperKind), SrcConstants.PrintPaperKind, this.PaperKind);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperSource), SrcConstants.PrintPaperSource, this.PaperSource);
            yield return new CellTuple(nameof(SrcConstants.PrintRightMargin), SrcConstants.PrintRightMargin, this.RightMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleX), SrcConstants.PrintScaleX, this.ScaleX);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleY), SrcConstants.PrintScaleY, this.ScaleY);
            yield return new CellTuple(nameof(SrcConstants.PrintTopMargin), SrcConstants.PrintTopMargin, this.TopMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintBottomMargin), SrcConstants.PrintBottomMargin, this.BottomMargin);
        }
    }
}