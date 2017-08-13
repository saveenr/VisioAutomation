using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class ShapeFormatCells : VisioPowerShell.Models.BaseCells
    {
        public string FillBackground;
        public string FillBackgroundTransparency;
        public string FillForeground;
        public string FillForegroundTransparency;
        public string FillPattern;
        public string FillShadowBackground;
        public string FillShadowBackgroundTransparency;
        public string FillShadowForeground;
        public string FillShadowForegroundTransparency;
        public string FillShadowPattern;
        public string GroupSelectMode;
        public string LineBeginArrow;
        public string LineBeginArrowSize;
        public string LineCap;
        public string LineColor;
        public string LineEndArrow;
        public string LineEndArrowSize;
        public string LinePattern;
        public string LineRounding;
        public string LineWeight;
        public string OneDBeginX;
        public string OneDBeginY;
        public string OneDEndX;
        public string OneDEndY;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.FillBackground), SRCCON.FillBackground, this.FillBackground);
            yield return new CellTuple(nameof(SRCCON.FillBackgroundTransparency), SRCCON.FillBackgroundTransparency, this.FillBackgroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillForeground), SRCCON.FillForeground, this.FillForeground);
            yield return new CellTuple(nameof(SRCCON.FillForegroundTransparency), SRCCON.FillForegroundTransparency, this.FillForegroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillPattern), SRCCON.FillPattern, this.FillPattern);
            yield return new CellTuple(nameof(SRCCON.FillShadowBackground), SRCCON.FillShadowBackground, this.FillShadowBackground);
            yield return new CellTuple(nameof(SRCCON.FillShadowBackgroundTransparency), SRCCON.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillShadowForeground), SRCCON.FillShadowForeground, this.FillShadowForeground);
            yield return new CellTuple(nameof(SRCCON.FillShadowForegroundTransparency), SRCCON.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillShadowPattern), SRCCON.FillShadowPattern, this.FillShadowPattern);
            yield return new CellTuple(nameof(SRCCON.GroupSelectMode), SRCCON.GroupSelectMode, this.GroupSelectMode);
            yield return new CellTuple(nameof(SRCCON.LineBeginArrow), SRCCON.LineBeginArrow, this.LineBeginArrow);
            yield return new CellTuple(nameof(SRCCON.LineBeginArrowSize), SRCCON.LineBeginArrowSize, this.LineBeginArrowSize);
            yield return new CellTuple(nameof(SRCCON.LineCap), SRCCON.LineCap, this.LineCap);
            yield return new CellTuple(nameof(SRCCON.LineColor), SRCCON.LineColor, this.LineColor);
            yield return new CellTuple(nameof(SRCCON.LineEndArrow), SRCCON.LineEndArrow, this.LineEndArrow);
            yield return new CellTuple(nameof(SRCCON.LineEndArrowSize), SRCCON.LineEndArrowSize, this.LineEndArrowSize);
            yield return new CellTuple(nameof(SRCCON.LinePattern), SRCCON.LinePattern, this.LinePattern);
            yield return new CellTuple(nameof(SRCCON.LineRounding), SRCCON.LineRounding, this.LineRounding);
            yield return new CellTuple(nameof(SRCCON.LineWeight), SRCCON.LineWeight, this.LineWeight);
            yield return new CellTuple(nameof(SRCCON.OneDBeginX), SRCCON.OneDBeginX, this.OneDBeginX);
            yield return new CellTuple(nameof(SRCCON.OneDBeginY), SRCCON.OneDBeginY, this.OneDBeginY);
            yield return new CellTuple(nameof(SRCCON.OneDEndX), SRCCON.OneDEndX, this.OneDEndX);
            yield return new CellTuple(nameof(SRCCON.OneDEndY), SRCCON.OneDEndY, this.OneDEndY);
        }
    }
}

