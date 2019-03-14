using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral DrawingScale { get; set; }
        public VASS.CellValueLiteral DrawingScaleType { get; set; }
        public VASS.CellValueLiteral DrawingSizeType { get; set; }
        public VASS.CellValueLiteral InhibitSnap { get; set; }
        public VASS.CellValueLiteral Height { get; set; }
        public VASS.CellValueLiteral Scale { get; set; }
        public VASS.CellValueLiteral Width { get; set; }
        public VASS.CellValueLiteral ShadowObliqueAngle { get; set; }
        public VASS.CellValueLiteral ShadowOffsetX { get; set; }
        public VASS.CellValueLiteral ShadowOffsetY { get; set; }
        public VASS.CellValueLiteral ShadowScaleFactor { get; set; }
        public VASS.CellValueLiteral ShadowType { get; set; }
        public VASS.CellValueLiteral UIVisibility { get; set; }
        public VASS.CellValueLiteral DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<VASS.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            { 
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingScale, this.DrawingScale);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageInhibitSnap, this.InhibitSnap);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageHeight, this.Height);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageScale, this.Scale);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageWidth, this.Width);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageShadowType, this.ShadowType);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageUIVisibility, this.UIVisibility);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            }
        }

    }
}