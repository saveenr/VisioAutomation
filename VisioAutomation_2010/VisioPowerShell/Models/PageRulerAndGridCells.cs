using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageRulerAndGridCells : VisioPowerShell.Models.BaseCells
    {
        public string XGridDensity;
        public string XGridOrigin;
        public string XGridSpacing;
        public string XRulerDensity;
        public string XRulerOrigin;
        public string YGridDensity;
        public string YGridOrigin;
        public string YGridSpacing;
        public string YRulerDensity;
        public string YRulerOrigin;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.XGridDensity), SrcConstants.XGridDensity, this.XGridDensity);
            yield return new CellTuple(nameof(SrcConstants.XGridOrigin), SrcConstants.XGridOrigin, this.XGridOrigin);
            yield return new CellTuple(nameof(SrcConstants.XGridSpacing), SrcConstants.XGridSpacing, this.XGridSpacing);
            yield return new CellTuple(nameof(SrcConstants.XRulerDensity), SrcConstants.XRulerDensity, this.XRulerDensity);
            yield return new CellTuple(nameof(SrcConstants.XRulerOrigin), SrcConstants.XRulerOrigin, this.XRulerOrigin);
            yield return new CellTuple(nameof(SrcConstants.YGridDensity), SrcConstants.YGridDensity, this.YGridDensity);
            yield return new CellTuple(nameof(SrcConstants.YGridOrigin), SrcConstants.YGridOrigin, this.YGridOrigin);
            yield return new CellTuple(nameof(SrcConstants.YGridSpacing), SrcConstants.YGridSpacing, this.YGridSpacing);
            yield return new CellTuple(nameof(SrcConstants.YRulerDensity), SrcConstants.YRulerDensity, this.YRulerDensity);
            yield return new CellTuple(nameof(SrcConstants.YRulerOrigin), SrcConstants.YRulerOrigin, this.YRulerOrigin);
        }
    }
}