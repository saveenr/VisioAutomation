using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PageRulerAndGridCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral XGridDensity { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YGridDensity { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XGridOrigin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YGridOrigin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XGridSpacing { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YGridSpacing { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XRulerDensity { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral XRulerOrigin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YRulerDensity { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral YRulerOrigin { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XGridDensity, this.XGridDensity.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XGridOrigin, this.XGridOrigin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XGridSpacing, this.XGridSpacing.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XRulerDensity, this.XRulerDensity.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.XRulerOrigin, this.XRulerOrigin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.YGridDensity, this.YGridDensity.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.YGridOrigin, this.YGridOrigin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.YGridSpacing, this.YGridSpacing.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.YRulerDensity, this.YRulerDensity.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.YRulerOrigin, this.YRulerOrigin.Value);
            }
        }

        public static PageRulerAndGridCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageRulerAndGridCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static PageRulerAndGridCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageRulerAndGridCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> lazy_query = new System.Lazy<PageRulerAndGridCellsReader>();

        class PageRulerAndGridCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageRulerAndGridCells>
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
                this.XGridDensity = this.query.Columns.Add(SrcConstants.XGridDensity, nameof(SrcConstants.XGridDensity));
                this.XGridOrigin = this.query.Columns.Add(SrcConstants.XGridOrigin, nameof(SrcConstants.XGridOrigin));
                this.XGridSpacing = this.query.Columns.Add(SrcConstants.XGridSpacing, nameof(SrcConstants.XGridSpacing));
                this.XRulerDensity = this.query.Columns.Add(SrcConstants.XRulerDensity, nameof(SrcConstants.XRulerDensity));
                this.XRulerOrigin = this.query.Columns.Add(SrcConstants.XRulerOrigin, nameof(SrcConstants.XRulerOrigin));
                this.YGridDensity = this.query.Columns.Add(SrcConstants.YGridDensity, nameof(SrcConstants.YGridDensity));
                this.YGridOrigin = this.query.Columns.Add(SrcConstants.YGridOrigin, nameof(SrcConstants.YGridOrigin));
                this.YGridSpacing = this.query.Columns.Add(SrcConstants.YGridSpacing, nameof(SrcConstants.YGridSpacing));
                this.YRulerDensity = this.query.Columns.Add(SrcConstants.YRulerDensity, nameof(SrcConstants.YRulerDensity));
                this.YRulerOrigin = this.query.Columns.Add(SrcConstants.YRulerOrigin, nameof(SrcConstants.YRulerOrigin));
            }

            public override PageRulerAndGridCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Pages.PageRulerAndGridCells();
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