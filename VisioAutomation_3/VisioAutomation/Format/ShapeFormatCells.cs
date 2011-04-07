using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public class ShapeFormatCells
    {
        // Fill
        public VA.ShapeSheet.CellData<int> FillBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double>FillBkgndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> FillForegnd { get; set; }
        public VA.ShapeSheet.CellData<double> FillForegndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> FillPattern { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwObliqueAngle { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwOffsetX { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwOffsetY { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwScaleFactor { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeShdwType { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwBkgndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwForegnd { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwForegndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwPattern { get; set; }

        // Line
        public VA.ShapeSheet.CellData<int> BeginArrow { get; set; }
        public VA.ShapeSheet.CellData<double> BeginArrowSize { get; set; }
        public VA.ShapeSheet.CellData<int> EndArrow { get; set; }
        public VA.ShapeSheet.CellData<double> EndArrowSize { get; set; }
        public VA.ShapeSheet.CellData<int> LineCap { get; set; }
        public VA.ShapeSheet.CellData<int> LineColor { get; set; }
        public VA.ShapeSheet.CellData<double> LineColorTrans { get; set; }
        public VA.ShapeSheet.CellData<int> LinePattern { get; set; }
        public VA.ShapeSheet.CellData<double> LineWeight { get; set; }
        public VA.ShapeSheet.CellData<double> Rounding { get; set; }

        // Char
        public VA.ShapeSheet.CellData<int> CharFont { get; set; }
        public VA.ShapeSheet.CellData<int> CharColor { get; set; }
        public VA.ShapeSheet.CellData<double> CharColorTrans { get; set; }
        public VA.ShapeSheet.CellData<double> CharSize { get; set; }

        // Text

        public VA.ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> TextBkgndTrans { get; set; }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        public void _Apply(System.Action<VA.ShapeSheet.SRC, VA.ShapeSheet.FormulaLiteral> func)
        {

            // Fill
            func(ShapeSheet.SRCConstants.FillBkgnd, FillBkgnd.Formula);
            func(ShapeSheet.SRCConstants.FillBkgndTrans, FillBkgndTrans.Formula);
            func(ShapeSheet.SRCConstants.FillForegnd, FillForegnd.Formula);
            func(ShapeSheet.SRCConstants.FillForegndTrans, FillForegndTrans.Formula);
            func(ShapeSheet.SRCConstants.FillPattern, FillPattern.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, ShapeShdwObliqueAngle.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwOffsetX, ShapeShdwOffsetX.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwOffsetY, ShapeShdwOffsetY.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, ShapeShdwScaleFactor.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwType, ShapeShdwType.Formula);
            func(ShapeSheet.SRCConstants.ShdwBkgnd, ShdwBkgnd.Formula);
            func(ShapeSheet.SRCConstants.ShdwBkgndTrans, ShdwBkgndTrans.Formula);
            func(ShapeSheet.SRCConstants.ShdwForegnd, ShdwForegnd.Formula);
            func(ShapeSheet.SRCConstants.ShdwForegndTrans, ShdwForegndTrans.Formula);
            func(ShapeSheet.SRCConstants.ShdwPattern, ShdwPattern.Formula);


            // Line
            func(ShapeSheet.SRCConstants.BeginArrow, BeginArrow.Formula);
            func(ShapeSheet.SRCConstants.BeginArrowSize, BeginArrowSize.Formula);
            func(ShapeSheet.SRCConstants.LineColor, LineColor.Formula);
            func(ShapeSheet.SRCConstants.LineColorTrans, LineColorTrans.Formula);
            func(ShapeSheet.SRCConstants.LinePattern, LinePattern.Formula);
            func(ShapeSheet.SRCConstants.LineWeight, LineWeight.Formula);
            func(ShapeSheet.SRCConstants.EndArrow, EndArrow.Formula);
            func(ShapeSheet.SRCConstants.EndArrowSize, EndArrowSize.Formula);

            // Char
            func(ShapeSheet.SRCConstants.Char_Color, CharColor.Formula);
            func(ShapeSheet.SRCConstants.Char_ColorTrans, CharColorTrans.Formula);
            func(ShapeSheet.SRCConstants.Char_Font, CharFont.Formula);
            func(ShapeSheet.SRCConstants.Char_Size, CharSize.Formula);

            // Text
            func(ShapeSheet.SRCConstants.TextBkgnd, TextBkgnd.Formula);
            func(ShapeSheet.SRCConstants.TextBkgndTrans, TextBkgndTrans.Formula);
        }
    }
}