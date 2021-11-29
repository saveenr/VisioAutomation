using System.Collections.Generic;
using VACG=VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PrintCells : VACG.CellGroup
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

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.LeftMargin), Core.SrcConstants.PrintLeftMargin, this.LeftMargin);
            yield return this._create(nameof(this.CenterX), Core.SrcConstants.PrintCenterX, this.CenterX);
            yield return this._create(nameof(this.CenterY), Core.SrcConstants.PrintCenterY, this.CenterY);
            yield return this._create(nameof(this.OnPage), Core.SrcConstants.PrintOnPage, this.OnPage);
            yield return this._create(nameof(this.BottomMargin), Core.SrcConstants.PrintBottomMargin, this.BottomMargin);
            yield return this._create(nameof(this.RightMargin), Core.SrcConstants.PrintRightMargin, this.RightMargin);
            yield return this._create(nameof(this.PagesX), Core.SrcConstants.PrintPagesX, this.PagesX);
            yield return this._create(nameof(this.PagesY), Core.SrcConstants.PrintPagesY, this.PagesY);
            yield return this._create(nameof(this.TopMargin), Core.SrcConstants.PrintTopMargin, this.TopMargin);
            yield return this._create(nameof(this.PaperKind), Core.SrcConstants.PrintPaperKind, this.PaperKind);
            yield return this._create(nameof(this.Grid), Core.SrcConstants.PrintGrid, this.Grid);
            yield return this._create(nameof(this.Orientation), Core.SrcConstants.PrintPageOrientation, this.Orientation);
            yield return this._create(nameof(this.ScaleX), Core.SrcConstants.PrintScaleX, this.ScaleX);
            yield return this._create(nameof(this.ScaleY), Core.SrcConstants.PrintScaleY, this.ScaleY);
            yield return this._create(nameof(this.PaperSource), Core.SrcConstants.PrintPaperSource, this.PaperSource);
        }


        public static PrintCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : VACG.CellGroupBuilder<PrintCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }

            public override PrintCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new PrintCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);


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