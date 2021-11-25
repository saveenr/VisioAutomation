using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue LeftMargin { get; set; }
        public VASS.CellValue CenterX { get; set; }
        public VASS.CellValue CenterY { get; set; }
        public VASS.CellValue OnPage { get; set; }
        public VASS.CellValue BottomMargin { get; set; }
        public VASS.CellValue RightMargin { get; set; }
        public VASS.CellValue PagesX { get; set; }
        public VASS.CellValue PagesY { get; set; }
        public VASS.CellValue TopMargin { get; set; }
        public VASS.CellValue PaperKind { get; set; }
        public VASS.CellValue Grid { get; set; }
        public VASS.CellValue Orientation { get; set; }
        public VASS.CellValue ScaleX { get; set; }
        public VASS.CellValue ScaleY { get; set; }
        public VASS.CellValue PaperSource { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.LeftMargin), VASS.SrcConstants.PrintLeftMargin, this.LeftMargin);
            yield return this.Create(nameof(this.CenterX), VASS.SrcConstants.PrintCenterX, this.CenterX);
            yield return this.Create(nameof(this.CenterY), VASS.SrcConstants.PrintCenterY, this.CenterY);
            yield return this.Create(nameof(this.OnPage), VASS.SrcConstants.PrintOnPage, this.OnPage);
            yield return this.Create(nameof(this.BottomMargin), VASS.SrcConstants.PrintBottomMargin, this.BottomMargin);
            yield return this.Create(nameof(this.RightMargin), VASS.SrcConstants.PrintRightMargin, this.RightMargin);
            yield return this.Create(nameof(this.PagesX), VASS.SrcConstants.PrintPagesX, this.PagesX);
            yield return this.Create(nameof(this.PagesY), VASS.SrcConstants.PrintPagesY, this.PagesY);
            yield return this.Create(nameof(this.TopMargin), VASS.SrcConstants.PrintTopMargin, this.TopMargin);
            yield return this.Create(nameof(this.PaperKind), VASS.SrcConstants.PrintPaperKind, this.PaperKind);
            yield return this.Create(nameof(this.Grid), VASS.SrcConstants.PrintGrid, this.Grid);
            yield return this.Create(nameof(this.Orientation), VASS.SrcConstants.PrintPageOrientation, this.Orientation);
            yield return this.Create(nameof(this.ScaleX), VASS.SrcConstants.PrintScaleX, this.ScaleX);
            yield return this.Create(nameof(this.ScaleY), VASS.SrcConstants.PrintScaleY, this.ScaleY);
            yield return this.Create(nameof(this.PaperSource), VASS.SrcConstants.PrintPaperSource, this.PaperSource);
        }


        public static PagePrintCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
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