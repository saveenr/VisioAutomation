using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;

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

        public override IEnumerable<VASS.CellGroups.NamedSrcValuePair> NamedSrcValuePairs
        {
            get
            {
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.XGridDensity), VASS.SrcConstants.XGridDensity, this.XGridDensity);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.XGridOrigin), VASS.SrcConstants.XGridOrigin, this.XGridOrigin);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.XGridSpacing), VASS.SrcConstants.XGridSpacing, this.XGridSpacing);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.XRulerDensity), VASS.SrcConstants.XRulerDensity, this.XRulerDensity);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.XRulerOrigin), VASS.SrcConstants.XRulerOrigin, this.XRulerOrigin);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.YGridDensity), VASS.SrcConstants.YGridDensity, this.YGridDensity);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.YGridOrigin), VASS.SrcConstants.YGridOrigin, this.YGridOrigin);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.YGridSpacing), VASS.SrcConstants.YGridSpacing, this.YGridSpacing);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.YRulerDensity), VASS.SrcConstants.YRulerDensity, this.YRulerDensity);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.YRulerOrigin), VASS.SrcConstants.YRulerOrigin, this.YRulerOrigin);
            }
        }
    }
}