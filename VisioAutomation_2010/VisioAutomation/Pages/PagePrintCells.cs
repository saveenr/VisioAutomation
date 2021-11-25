using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue LeftMargin { get; set; }
        public VisioAutomation.Core.CellValue CenterX { get; set; }
        public VisioAutomation.Core.CellValue CenterY { get; set; }
        public VisioAutomation.Core.CellValue OnPage { get; set; }
        public VisioAutomation.Core.CellValue BottomMargin { get; set; }
        public VisioAutomation.Core.CellValue RightMargin { get; set; }
        public VisioAutomation.Core.CellValue PagesX { get; set; }
        public VisioAutomation.Core.CellValue PagesY { get; set; }
        public VisioAutomation.Core.CellValue TopMargin { get; set; }
        public VisioAutomation.Core.CellValue PaperKind { get; set; }
        public VisioAutomation.Core.CellValue Grid { get; set; }
        public VisioAutomation.Core.CellValue Orientation { get; set; }
        public VisioAutomation.Core.CellValue ScaleX { get; set; }
        public VisioAutomation.Core.CellValue ScaleY { get; set; }
        public VisioAutomation.Core.CellValue PaperSource { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.LeftMargin), VisioAutomation.Core.SrcConstants.PrintLeftMargin, this.LeftMargin);
            yield return this.Create(nameof(this.CenterX), VisioAutomation.Core.SrcConstants.PrintCenterX, this.CenterX);
            yield return this.Create(nameof(this.CenterY), VisioAutomation.Core.SrcConstants.PrintCenterY, this.CenterY);
            yield return this.Create(nameof(this.OnPage), VisioAutomation.Core.SrcConstants.PrintOnPage, this.OnPage);
            yield return this.Create(nameof(this.BottomMargin), VisioAutomation.Core.SrcConstants.PrintBottomMargin, this.BottomMargin);
            yield return this.Create(nameof(this.RightMargin), VisioAutomation.Core.SrcConstants.PrintRightMargin, this.RightMargin);
            yield return this.Create(nameof(this.PagesX), VisioAutomation.Core.SrcConstants.PrintPagesX, this.PagesX);
            yield return this.Create(nameof(this.PagesY), VisioAutomation.Core.SrcConstants.PrintPagesY, this.PagesY);
            yield return this.Create(nameof(this.TopMargin), VisioAutomation.Core.SrcConstants.PrintTopMargin, this.TopMargin);
            yield return this.Create(nameof(this.PaperKind), VisioAutomation.Core.SrcConstants.PrintPaperKind, this.PaperKind);
            yield return this.Create(nameof(this.Grid), VisioAutomation.Core.SrcConstants.PrintGrid, this.Grid);
            yield return this.Create(nameof(this.Orientation), VisioAutomation.Core.SrcConstants.PrintPageOrientation, this.Orientation);
            yield return this.Create(nameof(this.ScaleX), VisioAutomation.Core.SrcConstants.PrintScaleX, this.ScaleX);
            yield return this.Create(nameof(this.ScaleY), VisioAutomation.Core.SrcConstants.PrintScaleY, this.ScaleY);
            yield return this.Create(nameof(this.PaperSource), VisioAutomation.Core.SrcConstants.PrintPaperSource, this.PaperSource);
        }


        public static PagePrintCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = PagePrintCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PagePrintCellsBuilder> PagePrintCells_lazy_builder = new System.Lazy<PagePrintCellsBuilder>();

        class PagePrintCellsBuilder : VASS.CellGroups.CellGroupBuilder<PagePrintCells>
        {
            public PagePrintCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override PagePrintCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new PagePrintCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);


                cells.LeftMargin = getcellvalue(nameof(PagePrintCells.LeftMargin));
                cells.CenterX = getcellvalue(nameof(PagePrintCells.CenterX));
                cells.CenterY = getcellvalue(nameof(PagePrintCells.CenterY));

                cells.OnPage = getcellvalue(nameof(PagePrintCells.OnPage));
                cells.BottomMargin = getcellvalue(nameof(PagePrintCells.BottomMargin));
                cells.RightMargin = getcellvalue(nameof(PagePrintCells.RightMargin));
                cells.PagesX = getcellvalue(nameof(PagePrintCells.PagesX));
                cells.PagesY = getcellvalue(nameof(PagePrintCells.PagesY));
                cells.TopMargin = getcellvalue(nameof(PagePrintCells.TopMargin));
                cells.PaperKind = getcellvalue(nameof(PagePrintCells.PaperKind));

                cells.Grid = getcellvalue(nameof(PagePrintCells.Grid));
                cells.Orientation = getcellvalue(nameof(PagePrintCells.Orientation));
                cells.ScaleX = getcellvalue(nameof(PagePrintCells.ScaleX));
                cells.ScaleY = getcellvalue(nameof(PagePrintCells.ScaleY));
                cells.PaperSource = getcellvalue(nameof(PagePrintCells.PaperSource));

                return cells;
            }
        }

    }
}