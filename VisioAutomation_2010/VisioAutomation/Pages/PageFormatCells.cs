using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;

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

        public override IEnumerable<VASS.CellGroups.NamedSrcValuePair> NamedSrcValuePairs
        {
            get
            {


                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.DrawingScale), VASS.SrcConstants.PageDrawingScale, this.DrawingScale);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.DrawingScaleType), VASS.SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.DrawingSizeType), VASS.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.InhibitSnap), VASS.SrcConstants.PageInhibitSnap, this.InhibitSnap);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.Height), VASS.SrcConstants.PageHeight, this.Height);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.Scale), VASS.SrcConstants.PageScale, this.Scale);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.Width), VASS.SrcConstants.PageWidth, this.Width);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.ShadowObliqueAngle), VASS.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.ShadowOffsetX), VASS.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.ShadowOffsetY), VASS.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.ShadowScaleFactor), VASS.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.ShadowType), VASS.SrcConstants.PageShadowType, this.ShadowType);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.UIVisibility), VASS.SrcConstants.PageUIVisibility, this.UIVisibility);
                yield return VASS.CellGroups.NamedSrcValuePair.Create(nameof(this.DrawingResizeType), VASS.SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            }
        }

    }
}