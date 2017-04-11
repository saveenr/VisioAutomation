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
                yield return this.newpair(ShapeSheet.SrcConstants.PrintLeftMargin, this.LeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterX, this.CenterX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterY, this.CenterY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintOnPage, this.OnPage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintBottomMargin, this.BottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintRightMargin, this.RightMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesX, this.PagesX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesY, this.PagesY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintTopMargin, this.TopMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperKind, this.PaperKind.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintGrid, this.Grid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPageOrientation, this.Orientation.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleX, this.ScaleX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleY, this.ScaleY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperSource, this.PaperSource.Formula);
            }
        }

        public static PagePrintCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PagePrintCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();
    }
}