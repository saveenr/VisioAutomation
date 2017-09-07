using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData LeftMargin { get; set; }
        public ShapeSheet.CellData CenterX { get; set; }
        public ShapeSheet.CellData CenterY { get; set; }
        public ShapeSheet.CellData OnPage { get; set; }
        public ShapeSheet.CellData BottomMargin { get; set; }
        public ShapeSheet.CellData RightMargin { get; set; }
        public ShapeSheet.CellData PagesX { get; set; }
        public ShapeSheet.CellData PagesY { get; set; }
        public ShapeSheet.CellData TopMargin { get; set; }
        public ShapeSheet.CellData PaperKind { get; set; }
        public ShapeSheet.CellData Grid { get; set; }
        public ShapeSheet.CellData Orientation { get; set; }
        public ShapeSheet.CellData ScaleX { get; set; }
        public ShapeSheet.CellData ScaleY { get; set; }
        public ShapeSheet.CellData PaperSource { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.PrintLeftMargin, this.LeftMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterX, this.CenterX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterY, this.CenterY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintOnPage, this.OnPage.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintBottomMargin, this.BottomMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintRightMargin, this.RightMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesX, this.PagesX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesY, this.PagesY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintTopMargin, this.TopMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperKind, this.PaperKind.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintGrid, this.Grid.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPageOrientation, this.Orientation.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleX, this.ScaleX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleY, this.ScaleY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperSource, this.PaperSource.Value);
            }
        }

        public static PagePrintCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = PagePrintCells.lazy_query.Value;
            return query.GetCellGroup(shape, cvt);
        }

        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();
    }
}