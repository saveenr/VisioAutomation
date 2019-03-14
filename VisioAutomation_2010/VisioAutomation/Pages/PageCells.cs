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
    }
}