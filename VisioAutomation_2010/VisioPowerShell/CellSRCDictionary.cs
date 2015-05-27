using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    public class CellSRCDictionary : CellNameDictionary<SRC>
    {
        private static CellSRCDictionary shape_cellmap;
        private static CellSRCDictionary page_cellmap;

        public static CellSRCDictionary GetCellMapForShapes()
        {
            if (CellSRCDictionary.shape_cellmap == null)
            {
                CellSRCDictionary.shape_cellmap = new CellSRCDictionary();






                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Angle)] = SRCConstants.Angle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginX)] = SRCConstants.BeginX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginY)] = SRCConstants.BeginY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrow)] = SRCConstants.BeginArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrowSize)] = SRCConstants.BeginArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharCase)] = SRCConstants.CharCase;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColor)] = SRCConstants.CharColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColorTrans)] = SRCConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFont)] = SRCConstants.CharFont;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFontScale)] = SRCConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharLetterspace)] = SRCConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharSize)] = SRCConstants.CharSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharStyle)] = SRCConstants.CharStyle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndX)] = SRCConstants.EndX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndY)] = SRCConstants.EndY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrow)] = SRCConstants.EndArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrowSize)] = SRCConstants.EndArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgnd)] = SRCConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgndTrans)] = SRCConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegnd)] = SRCConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegndTrans)] = SRCConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillPattern)] = SRCConstants.FillPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Height)] = SRCConstants.Height;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineCap)] = SRCConstants.LineCap;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineColor)] = SRCConstants.LineColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LinePattern)] = SRCConstants.LinePattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineWeight)] = SRCConstants.LineWeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockAspect)] = SRCConstants.LockAspect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockBegin)] = SRCConstants.LockBegin;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCalcWH)] = SRCConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCrop)] = SRCConstants.LockCrop;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCustProp)] = SRCConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockDelete)] = SRCConstants.LockDelete;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockEnd)] = SRCConstants.LockEnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFormat)] = SRCConstants.LockFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFromGroupFormat)] = SRCConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockGroup)] = SRCConstants.LockGroup;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockHeight)] = SRCConstants.LockHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveX)] = SRCConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveY)] = SRCConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockRotate)] = SRCConstants.LockRotate;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockSelect)] = SRCConstants.LockSelect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockTextEdit)] = SRCConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeColors)] = SRCConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeEffects)] = SRCConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockVtxEdit)] = SRCConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockWidth)] = SRCConstants.LockWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinX)] = SRCConstants.LocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinY)] = SRCConstants.LocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinX)] = SRCConstants.PinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinY)] = SRCConstants.PinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Rounding)] = SRCConstants.Rounding;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.SelectMode)] = SRCConstants.SelectMode;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgnd)] = SRCConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgndTrans)] = SRCConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegnd)] = SRCConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegndTrans)] = SRCConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwObliqueAngle)] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetX)] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetY)] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwPattern)] = SRCConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwScaleFactor)] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwType)] = SRCConstants.ShdwType;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtAngle)] = SRCConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtHeight)] = SRCConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinX)] = SRCConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinY)] = SRCConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinX)] = SRCConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinY)] = SRCConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtWidth)] = SRCConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Width)] = SRCConstants.Width;

            }
            return CellSRCDictionary.shape_cellmap;
        }

        public static CellSRCDictionary GetCellMapForPages()
        {
            if (CellSRCDictionary.page_cellmap == null)
            {
                CellSRCDictionary.page_cellmap = new CellSRCDictionary();
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Angle)] = SRCConstants.Angle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginX)] = SRCConstants.BeginX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginY)] = SRCConstants.BeginY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrow)] = SRCConstants.BeginArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrowSize)] = SRCConstants.BeginArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharCase)] = SRCConstants.CharCase;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColor)] = SRCConstants.CharColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColorTrans)] = SRCConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFont)] = SRCConstants.CharFont;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFontScale)] = SRCConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharLetterspace)] = SRCConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharSize)] = SRCConstants.CharSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharStyle)] = SRCConstants.CharStyle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndX)] = SRCConstants.EndX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndY)] = SRCConstants.EndY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrow)] = SRCConstants.EndArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrowSize)] = SRCConstants.EndArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgnd)] = SRCConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgndTrans)] = SRCConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegnd)] = SRCConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegndTrans)] = SRCConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillPattern)] = SRCConstants.FillPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Height)] = SRCConstants.Height;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineCap)] = SRCConstants.LineCap;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineColor)] = SRCConstants.LineColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LinePattern)] = SRCConstants.LinePattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineWeight)] = SRCConstants.LineWeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockAspect)] = SRCConstants.LockAspect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockBegin)] = SRCConstants.LockBegin;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCalcWH)] = SRCConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCrop)] = SRCConstants.LockCrop;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCustProp)] = SRCConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockDelete)] = SRCConstants.LockDelete;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockEnd)] = SRCConstants.LockEnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFormat)] = SRCConstants.LockFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFromGroupFormat)] = SRCConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockGroup)] = SRCConstants.LockGroup;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockHeight)] = SRCConstants.LockHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveX)] = SRCConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveY)] = SRCConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockRotate)] = SRCConstants.LockRotate;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockSelect)] = SRCConstants.LockSelect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockTextEdit)] = SRCConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeColors)] = SRCConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeEffects)] = SRCConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockVtxEdit)] = SRCConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockWidth)] = SRCConstants.LockWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinX)] = SRCConstants.LocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinY)] = SRCConstants.LocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinX)] = SRCConstants.PinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinY)] = SRCConstants.PinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Rounding)] = SRCConstants.Rounding;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.SelectMode)] = SRCConstants.SelectMode;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgnd)] = SRCConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgndTrans)] = SRCConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegnd)] = SRCConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegndTrans)] = SRCConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwObliqueAngle)] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetX)] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetY)] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwPattern)] = SRCConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwScaleFactor)] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwType)] = SRCConstants.ShdwType;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtAngle)] = SRCConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtHeight)] = SRCConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinX)] = SRCConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinY)] = SRCConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinX)] = SRCConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinY)] = SRCConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtWidth)] = SRCConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Width)] = SRCConstants.Width;



            }
            return CellSRCDictionary.page_cellmap;
        }

        public CellQuery CreateQueryFromCellNames(IList<string> Cells)
        {
            var invalid_names = Cells.Where(cellname => !this.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new CellQuery();

            foreach (string resolved_cellname in this.ResolveNames(Cells))
            {
                if (!query.CellColumns.Contains(resolved_cellname))
                {
                    var resolved_src = this[resolved_cellname];
                    query.AddCell(resolved_src, resolved_cellname);
                }
            }
            return query;
        }
    }
}