using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral LeftMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CenterX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CenterY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral OnPage { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BottomMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RightMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PagesX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PagesY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TopMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PaperKind { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Grid { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Orientation { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ScaleX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ScaleY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PaperSource { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintLeftMargin, this.LeftMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintCenterX, this.CenterX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintCenterY, this.CenterY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintOnPage, this.OnPage.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintBottomMargin, this.BottomMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintRightMargin, this.RightMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPagesX, this.PagesX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPagesY, this.PagesY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintTopMargin, this.TopMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPaperKind, this.PaperKind.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintGrid, this.Grid.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPageOrientation, this.Orientation.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintScaleX, this.ScaleX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintScaleY, this.ScaleY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPaperSource, this.PaperSource.Value);
            }
        }

        public static PagePrintCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PagePrintCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static PagePrintCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PagePrintCells.lazy_query.Value;
            return query.GetResults(shape);
        }
        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();
    }
}