using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.Scripting
{
    class LockCells
    {
        public ShapeSheet.CellData<bool> LockAspect { get; set; }
        public ShapeSheet.CellData<bool> LockBegin { get; set; }
        public ShapeSheet.CellData<bool> LockCalcWH { get; set; }
        public ShapeSheet.CellData<bool> LockCrop { get; set; }
        public ShapeSheet.CellData<bool> LockCustProp { get; set; }
        public ShapeSheet.CellData<bool> LockDelete { get; set; }
        public ShapeSheet.CellData<bool> LockEnd { get; set; }
        public ShapeSheet.CellData<bool> LockFormat { get; set; }
        public ShapeSheet.CellData<bool> LockFromGroupFormat { get; set; }
        public ShapeSheet.CellData<bool> LockGroup { get; set; }
        public ShapeSheet.CellData<bool> LockHeight { get; set; }
        public ShapeSheet.CellData<bool> LockMoveX { get; set; }
        public ShapeSheet.CellData<bool> LockMoveY { get; set; }
        public ShapeSheet.CellData<bool> LockRotate { get; set; }
        public ShapeSheet.CellData<bool> LockSelect { get; set; }
        public ShapeSheet.CellData<bool> LockTextEdit { get; set; }
        public ShapeSheet.CellData<bool> LockThemeColors { get; set; }
        public ShapeSheet.CellData<bool> LockThemeEffects { get; set; }
        public ShapeSheet.CellData<bool> LockVtxEdit { get; set; }
        public ShapeSheet.CellData<bool> LockWidth { get; set; }

        public void Apply(ShapeSheet.Update.SIDSRCUpdate update, short id)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f));
        }

        public void Apply(ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        internal void _Apply(System.Action<ShapeSheet.SRC, ShapeSheet.FormulaLiteral> func)
        {

            func(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
            func(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
            func(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
            func(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
            func(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
            func(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
            func(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
            func(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
            func(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
            func(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
            func(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
            func(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
            func(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
            func(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
            func(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
            func(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
            func(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
            func(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
            func(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
            func(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
        }

        public void SetAll(string formula)
        {
            LockAspect = formula;
            LockBegin = formula;
            LockCalcWH = formula;
            LockCrop = formula;
            LockCustProp = formula;
            LockDelete = formula;
            LockEnd = formula;
            LockFormat = formula;
            LockFromGroupFormat = formula;
            LockGroup = formula;
            LockHeight = formula;
            LockMoveX = formula;
            LockMoveY = formula;
            LockRotate = formula;
            LockSelect = formula;
            LockTextEdit = formula;
            LockThemeColors = formula;
            LockThemeEffects = formula;
            LockVtxEdit = formula;
            LockWidth = formula;
        }

    }
}