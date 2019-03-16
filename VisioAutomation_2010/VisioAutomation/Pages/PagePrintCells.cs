using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

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

            public override PagePrintCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new PagePrintCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }


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