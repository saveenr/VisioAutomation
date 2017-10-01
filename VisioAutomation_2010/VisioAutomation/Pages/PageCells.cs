using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PageRulerAndGridCells : CellGroupSingleRow
    {
        public CellValueLiteral XGridDensity { get; set; }
        public CellValueLiteral YGridDensity { get; set; }
        public CellValueLiteral XGridOrigin { get; set; }
        public CellValueLiteral YGridOrigin { get; set; }
        public CellValueLiteral XGridSpacing { get; set; }
        public CellValueLiteral YGridSpacing { get; set; }
        public CellValueLiteral XRulerDensity { get; set; }
        public CellValueLiteral XRulerOrigin { get; set; }
        public CellValueLiteral YRulerDensity { get; set; }
        public CellValueLiteral YRulerOrigin { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.XGridDensity, this.XGridDensity);
                yield return SrcValuePair.Create(SrcConstants.XGridOrigin, this.XGridOrigin);
                yield return SrcValuePair.Create(SrcConstants.XGridSpacing, this.XGridSpacing);
                yield return SrcValuePair.Create(SrcConstants.XRulerDensity, this.XRulerDensity);
                yield return SrcValuePair.Create(SrcConstants.XRulerOrigin, this.XRulerOrigin);
                yield return SrcValuePair.Create(SrcConstants.YGridDensity, this.YGridDensity);
                yield return SrcValuePair.Create(SrcConstants.YGridOrigin, this.YGridOrigin);
                yield return SrcValuePair.Create(SrcConstants.YGridSpacing, this.YGridSpacing);
                yield return SrcValuePair.Create(SrcConstants.YRulerDensity, this.YRulerDensity);
                yield return SrcValuePair.Create(SrcConstants.YRulerOrigin, this.YRulerOrigin);
            }
        }

        public static PageRulerAndGridCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> lazy_query = new System.Lazy<PageRulerAndGridCellsReader>();

        class PageRulerAndGridCellsReader : ReaderSingleRow<PageRulerAndGridCells>
        {
            public CellColumn XGridDensity { get; set; }
            public CellColumn XGridOrigin { get; set; }
            public CellColumn XGridSpacing { get; set; }
            public CellColumn XRulerDensity { get; set; }
            public CellColumn XRulerOrigin { get; set; }
            public CellColumn YGridDensity { get; set; }
            public CellColumn YGridOrigin { get; set; }
            public CellColumn YGridSpacing { get; set; }
            public CellColumn YRulerDensity { get; set; }
            public CellColumn YRulerOrigin { get; set; }

            public PageRulerAndGridCellsReader()
            {
                this.XGridDensity = this.query.Columns.Add(SrcConstants.XGridDensity, nameof(this.XGridDensity));
                this.XGridOrigin = this.query.Columns.Add(SrcConstants.XGridOrigin, nameof(this.XGridOrigin));
                this.XGridSpacing = this.query.Columns.Add(SrcConstants.XGridSpacing, nameof(this.XGridSpacing));
                this.XRulerDensity = this.query.Columns.Add(SrcConstants.XRulerDensity, nameof(this.XRulerDensity));
                this.XRulerOrigin = this.query.Columns.Add(SrcConstants.XRulerOrigin, nameof(this.XRulerOrigin));
                this.YGridDensity = this.query.Columns.Add(SrcConstants.YGridDensity, nameof(this.YGridDensity));
                this.YGridOrigin = this.query.Columns.Add(SrcConstants.YGridOrigin, nameof(this.YGridOrigin));
                this.YGridSpacing = this.query.Columns.Add(SrcConstants.YGridSpacing, nameof(this.YGridSpacing));
                this.YRulerDensity = this.query.Columns.Add(SrcConstants.YRulerDensity, nameof(this.YRulerDensity));
                this.YRulerOrigin = this.query.Columns.Add(SrcConstants.YRulerOrigin, nameof(this.YRulerOrigin));
            }

            public override PageRulerAndGridCells ToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new PageRulerAndGridCells();
                cells.XGridDensity = row[this.XGridDensity];
                cells.XGridOrigin = row[this.XGridOrigin];
                cells.XGridSpacing = row[this.XGridSpacing];
                cells.XRulerDensity = row[this.XRulerDensity];
                cells.XRulerOrigin = row[this.XRulerOrigin];
                cells.YGridDensity = row[this.YGridDensity];
                cells.YGridOrigin = row[this.YGridOrigin];
                cells.YGridSpacing = row[this.YGridSpacing];
                cells.YRulerDensity = row[this.YRulerDensity];
                cells.YRulerOrigin = row[this.YRulerOrigin];
                return cells;
            }
        }

    }
}