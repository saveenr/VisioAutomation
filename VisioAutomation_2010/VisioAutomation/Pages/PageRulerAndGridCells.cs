using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PageRulerAndGridCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData XGridDensity { get; set; }
        public ShapeSheet.CellData YGridDensity { get; set; }

        public ShapeSheet.CellData XGridOrigin { get; set; }
        public ShapeSheet.CellData YGridOrigin { get; set; }

        public ShapeSheet.CellData XGridSpacing { get; set; }
        public ShapeSheet.CellData YGridSpacing { get; set; }

        public ShapeSheet.CellData XRulerDensity { get; set; }
        public ShapeSheet.CellData XRulerOrigin { get; set; }

        public ShapeSheet.CellData YRulerDensity { get; set; }
        public ShapeSheet.CellData YRulerOrigin { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.XGridDensity, this.XGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XGridOrigin, this.XGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XGridSpacing, this.XGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XRulerDensity, this.XRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XRulerOrigin, this.XRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridDensity, this.YGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridOrigin, this.YGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridSpacing, this.YGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YRulerDensity, this.YRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YRulerOrigin, this.YRulerOrigin.Formula);
            }
        }

        public static PageRulerAndGridCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageRulerAndGridCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PageRulerAndGridCellsReader> lazy_query = new System.Lazy<PageRulerAndGridCellsReader>();
    }
}