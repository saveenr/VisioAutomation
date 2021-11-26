using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PrintCells : CellGroup
    {
        public Core.CellValue LeftMargin { get; set; }
        public Core.CellValue CenterX { get; set; }
        public Core.CellValue CenterY { get; set; }
        public Core.CellValue OnPage { get; set; }
        public Core.CellValue BottomMargin { get; set; }
        public Core.CellValue RightMargin { get; set; }
        public Core.CellValue PagesX { get; set; }
        public Core.CellValue PagesY { get; set; }
        public Core.CellValue TopMargin { get; set; }
        public Core.CellValue PaperKind { get; set; }
        public Core.CellValue Grid { get; set; }
        public Core.CellValue Orientation { get; set; }
        public Core.CellValue ScaleX { get; set; }
        public Core.CellValue ScaleY { get; set; }
        public Core.CellValue PaperSource { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.LeftMargin), Core.SrcConstants.PrintLeftMargin, this.LeftMargin);
            yield return this.Create(nameof(this.CenterX), Core.SrcConstants.PrintCenterX, this.CenterX);
            yield return this.Create(nameof(this.CenterY), Core.SrcConstants.PrintCenterY, this.CenterY);
            yield return this.Create(nameof(this.OnPage), Core.SrcConstants.PrintOnPage, this.OnPage);
            yield return this.Create(nameof(this.BottomMargin), Core.SrcConstants.PrintBottomMargin, this.BottomMargin);
            yield return this.Create(nameof(this.RightMargin), Core.SrcConstants.PrintRightMargin, this.RightMargin);
            yield return this.Create(nameof(this.PagesX), Core.SrcConstants.PrintPagesX, this.PagesX);
            yield return this.Create(nameof(this.PagesY), Core.SrcConstants.PrintPagesY, this.PagesY);
            yield return this.Create(nameof(this.TopMargin), Core.SrcConstants.PrintTopMargin, this.TopMargin);
            yield return this.Create(nameof(this.PaperKind), Core.SrcConstants.PrintPaperKind, this.PaperKind);
            yield return this.Create(nameof(this.Grid), Core.SrcConstants.PrintGrid, this.Grid);
            yield return this.Create(nameof(this.Orientation), Core.SrcConstants.PrintPageOrientation, this.Orientation);
            yield return this.Create(nameof(this.ScaleX), Core.SrcConstants.PrintScaleX, this.ScaleX);
            yield return this.Create(nameof(this.ScaleY), Core.SrcConstants.PrintScaleY, this.ScaleY);
            yield return this.Create(nameof(this.PaperSource), Core.SrcConstants.PrintPaperSource, this.PaperSource);
        }


        public static PrintCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = PagePrintCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PagePrintCellsBuilder> PagePrintCells_lazy_builder = new System.Lazy<PagePrintCellsBuilder>();

        class PagePrintCellsBuilder : CellGroupBuilder<PrintCells>
        {
            public PagePrintCellsBuilder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override PrintCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new PrintCells();
                var getcellvalue = row_to_cellgroup(row, cols);


                cells.LeftMargin = getcellvalue(nameof(LeftMargin));
                cells.CenterX = getcellvalue(nameof(CenterX));
                cells.CenterY = getcellvalue(nameof(CenterY));

                cells.OnPage = getcellvalue(nameof(OnPage));
                cells.BottomMargin = getcellvalue(nameof(BottomMargin));
                cells.RightMargin = getcellvalue(nameof(RightMargin));
                cells.PagesX = getcellvalue(nameof(PagesX));
                cells.PagesY = getcellvalue(nameof(PagesY));
                cells.TopMargin = getcellvalue(nameof(TopMargin));
                cells.PaperKind = getcellvalue(nameof(PaperKind));

                cells.Grid = getcellvalue(nameof(Grid));
                cells.Orientation = getcellvalue(nameof(Orientation));
                cells.ScaleX = getcellvalue(nameof(ScaleX));
                cells.ScaleY = getcellvalue(nameof(ScaleY));
                cells.PaperSource = getcellvalue(nameof(PaperSource));

                return cells;
            }
        }

    }
}