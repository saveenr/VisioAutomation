using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

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
            return query.GetFormulas(shape);
        }

        public static PageRulerAndGridCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageRulerAndGridCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> lazy_query = new System.Lazy<PageRulerAndGridCellsReader>();
    }
}