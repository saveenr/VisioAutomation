using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class RulerAndGridCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue XGridDensity { get; set; }
        public VisioAutomation.Core.CellValue YGridDensity { get; set; }
        public VisioAutomation.Core.CellValue XGridOrigin { get; set; }
        public VisioAutomation.Core.CellValue YGridOrigin { get; set; }
        public VisioAutomation.Core.CellValue XGridSpacing { get; set; }
        public VisioAutomation.Core.CellValue YGridSpacing { get; set; }
        public VisioAutomation.Core.CellValue XRulerDensity { get; set; }
        public VisioAutomation.Core.CellValue XRulerOrigin { get; set; }
        public VisioAutomation.Core.CellValue YRulerDensity { get; set; }
        public VisioAutomation.Core.CellValue YRulerOrigin { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.XGridDensity), VisioAutomation.Core.SrcConstants.XGridDensity, this.XGridDensity);
            yield return this.Create(nameof(this.XGridOrigin), VisioAutomation.Core.SrcConstants.XGridOrigin, this.XGridOrigin);
            yield return this.Create(nameof(this.XGridSpacing), VisioAutomation.Core.SrcConstants.XGridSpacing, this.XGridSpacing);
            yield return this.Create(nameof(this.XRulerDensity), VisioAutomation.Core.SrcConstants.XRulerDensity, this.XRulerDensity);
            yield return this.Create(nameof(this.XRulerOrigin), VisioAutomation.Core.SrcConstants.XRulerOrigin, this.XRulerOrigin);
            yield return this.Create(nameof(this.YGridDensity), VisioAutomation.Core.SrcConstants.YGridDensity, this.YGridDensity);
            yield return this.Create(nameof(this.YGridOrigin), VisioAutomation.Core.SrcConstants.YGridOrigin, this.YGridOrigin);
            yield return this.Create(nameof(this.YGridSpacing), VisioAutomation.Core.SrcConstants.YGridSpacing, this.YGridSpacing);
            yield return this.Create(nameof(this.YRulerDensity), VisioAutomation.Core.SrcConstants.YRulerDensity, this.YRulerDensity);
            yield return this.Create(nameof(this.YRulerOrigin), VisioAutomation.Core.SrcConstants.YRulerOrigin, this.YRulerOrigin);
        }

        public static RulerAndGridCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = PageRulerAndGridCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsBuilder> PageRulerAndGridCells_lazy_builder = new System.Lazy<PageRulerAndGridCellsBuilder>();

        class PageRulerAndGridCellsBuilder : VASS.CellGroups.CellGroupBuilder<RulerAndGridCells>
        {
            public PageRulerAndGridCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override RulerAndGridCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new RulerAndGridCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.XGridDensity = getcellvalue(nameof(RulerAndGridCells.XGridDensity));
                cells.XGridOrigin = getcellvalue(nameof(RulerAndGridCells.XGridOrigin));
                cells.XGridSpacing = getcellvalue(nameof(RulerAndGridCells.XGridSpacing));
                cells.XRulerDensity = getcellvalue(nameof(RulerAndGridCells.XRulerDensity));
                cells.XRulerOrigin = getcellvalue(nameof(RulerAndGridCells.XRulerOrigin));
                cells.YGridDensity = getcellvalue(nameof(RulerAndGridCells.YGridDensity));
                cells.YGridOrigin = getcellvalue(nameof(RulerAndGridCells.YGridOrigin));
                cells.YGridSpacing = getcellvalue(nameof(RulerAndGridCells.YGridSpacing));
                cells.YRulerDensity = getcellvalue(nameof(RulerAndGridCells.YRulerDensity));
                cells.YRulerOrigin = getcellvalue(nameof(RulerAndGridCells.YRulerOrigin));

                return cells;
            }
        }

    }
}