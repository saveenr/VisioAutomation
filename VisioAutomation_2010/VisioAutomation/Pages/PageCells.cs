using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageRulerAndGridCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral XGridDensity { get; set; }
        public VASS.CellValueLiteral YGridDensity { get; set; }
        public VASS.CellValueLiteral XGridOrigin { get; set; }
        public VASS.CellValueLiteral YGridOrigin { get; set; }
        public VASS.CellValueLiteral XGridSpacing { get; set; }
        public VASS.CellValueLiteral YGridSpacing { get; set; }
        public VASS.CellValueLiteral XRulerDensity { get; set; }
        public VASS.CellValueLiteral XRulerOrigin { get; set; }
        public VASS.CellValueLiteral YRulerDensity { get; set; }
        public VASS.CellValueLiteral YRulerOrigin { get; set; }

        public override IEnumerable<VASS.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.XGridDensity, this.XGridDensity);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.XGridOrigin, this.XGridOrigin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.XGridSpacing, this.XGridSpacing);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.XRulerDensity, this.XRulerDensity);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.XRulerOrigin, this.XRulerOrigin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.YGridDensity, this.YGridDensity);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.YGridOrigin, this.YGridOrigin);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.YGridSpacing, this.YGridSpacing);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.YRulerDensity, this.YRulerDensity);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.YRulerOrigin, this.YRulerOrigin);
            }
        }

        public static PageRulerAndGridCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> lazy_reader = new System.Lazy<PageRulerAndGridCellsReader>();

        class PageRulerAndGridCellsReader : VASS.CellGroups.CellGroupReader<PageRulerAndGridCells>
        {
            public VASS.Query.CellColumn XGridDensity { get; set; }
            public VASS.Query.CellColumn XGridOrigin { get; set; }
            public VASS.Query.CellColumn XGridSpacing { get; set; }
            public VASS.Query.CellColumn XRulerDensity { get; set; }
            public VASS.Query.CellColumn XRulerOrigin { get; set; }
            public VASS.Query.CellColumn YGridDensity { get; set; }
            public VASS.Query.CellColumn YGridOrigin { get; set; }
            public VASS.Query.CellColumn YGridSpacing { get; set; }
            public VASS.Query.CellColumn YRulerDensity { get; set; }
            public VASS.Query.CellColumn YRulerOrigin { get; set; }

            public PageRulerAndGridCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.XGridDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridDensity, nameof(this.XGridDensity));
                this.XGridOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridOrigin, nameof(this.XGridOrigin));
                this.XGridSpacing = this.query_singlerow.Columns.Add(VASS.SrcConstants.XGridSpacing, nameof(this.XGridSpacing));
                this.XRulerDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.XRulerDensity, nameof(this.XRulerDensity));
                this.XRulerOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.XRulerOrigin, nameof(this.XRulerOrigin));
                this.YGridDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridDensity, nameof(this.YGridDensity));
                this.YGridOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridOrigin, nameof(this.YGridOrigin));
                this.YGridSpacing = this.query_singlerow.Columns.Add(VASS.SrcConstants.YGridSpacing, nameof(this.YGridSpacing));
                this.YRulerDensity = this.query_singlerow.Columns.Add(VASS.SrcConstants.YRulerDensity, nameof(this.YRulerDensity));
                this.YRulerOrigin = this.query_singlerow.Columns.Add(VASS.SrcConstants.YRulerOrigin, nameof(this.YRulerOrigin));
            }

            public override PageRulerAndGridCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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