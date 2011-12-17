using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.Layout
{
    public partial class LockCells
    {
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