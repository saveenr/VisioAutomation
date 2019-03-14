using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : VASS.CellGroups.CellGroupBase
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

        public override IEnumerable<VASS.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintLeftMargin, this.LeftMargin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintCenterX, this.CenterX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintCenterY, this.CenterY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintOnPage, this.OnPage);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintBottomMargin, this.BottomMargin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintRightMargin, this.RightMargin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintPagesX, this.PagesX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintPagesY, this.PagesY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintTopMargin, this.TopMargin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintPaperKind, this.PaperKind);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintGrid, this.Grid);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintPageOrientation, this.Orientation);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintScaleX, this.ScaleX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintScaleY, this.ScaleY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PrintPaperSource, this.PaperSource);
            }
        }

        public static PagePrintCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();

        class PagePrintCellsReader : VASS.CellGroups.CellGroupReader<PagePrintCells>
        {
            public VASS.Query.CellColumn LeftMargin { get; set; }
            public VASS.Query.CellColumn CenterX { get; set; }
            public VASS.Query.CellColumn CenterY { get; set; }
            public VASS.Query.CellColumn OnPage { get; set; }
            public VASS.Query.CellColumn BottomMargin { get; set; }
            public VASS.Query.CellColumn RightMargin { get; set; }
            public VASS.Query.CellColumn PagesX { get; set; }
            public VASS.Query.CellColumn PagesY { get; set; }
            public VASS.Query.CellColumn TopMargin { get; set; }
            public VASS.Query.CellColumn PaperKind { get; set; }
            public VASS.Query.CellColumn Grid { get; set; }
            public VASS.Query.CellColumn PageOrientation { get; set; }
            public VASS.Query.CellColumn ScaleX { get; set; }
            public VASS.Query.CellColumn ScaleY { get; set; }
            public VASS.Query.CellColumn PaperSource { get; set; }

            public PagePrintCellsReader()
            {
                this.LeftMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintLeftMargin, nameof(this.LeftMargin));
                this.CenterX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintCenterX, nameof(this.CenterX));
                this.CenterY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintCenterY, nameof(this.CenterY));
                this.OnPage = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintOnPage, nameof(this.OnPage));
                this.BottomMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintBottomMargin, nameof(this.BottomMargin));
                this.RightMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintRightMargin, nameof(this.RightMargin));
                this.PagesX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPagesX, nameof(this.PagesX));
                this.PagesY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPagesY, nameof(this.PagesY));
                this.TopMargin = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintTopMargin, nameof(this.TopMargin));
                this.PaperKind = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPaperKind, nameof(this.PaperKind));
                this.Grid = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintGrid, nameof(this.Grid));
                this.PageOrientation = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPageOrientation, nameof(this.PageOrientation));
                this.ScaleX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintScaleX, nameof(this.ScaleX));
                this.ScaleY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintScaleY, nameof(this.ScaleY));
                this.PaperSource = this.query_singlerow.Columns.Add(VASS.SrcConstants.PrintPaperSource, nameof(this.PaperSource));
            }

            public override PagePrintCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PagePrintCells();
                cells.LeftMargin = row[this.LeftMargin];
                cells.CenterX = row[this.CenterX];
                cells.CenterY = row[this.CenterY];
                cells.OnPage = row[this.OnPage];
                cells.BottomMargin = row[this.BottomMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.PagesX = row[this.PagesX];
                cells.PagesY = row[this.PagesY];
                cells.TopMargin = row[this.TopMargin];
                cells.PaperKind = row[this.PaperKind];
                cells.Grid = row[this.Grid];
                cells.Orientation = row[this.PageOrientation];
                cells.ScaleX = row[this.ScaleX];
                cells.ScaleY = row[this.ScaleY];
                cells.PaperSource = row[this.PaperSource];
                return cells;
            }
        }

    }
}