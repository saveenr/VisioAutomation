using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.Layout
{
    public partial class LockCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<bool> LockAspect { get; set; }
        public VA.ShapeSheet.CellData<bool> LockBegin { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCalcWH { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCrop { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCustProp { get; set; }
        public VA.ShapeSheet.CellData<bool> LockDelete { get; set; }
        public VA.ShapeSheet.CellData<bool> LockEnd { get; set; }
        public VA.ShapeSheet.CellData<bool> LockFormat { get; set; }
        public VA.ShapeSheet.CellData<bool> LockFromGroupFormat { get; set; }
        public VA.ShapeSheet.CellData<bool> LockGroup { get; set; }
        public VA.ShapeSheet.CellData<bool> LockHeight { get; set; }
        public VA.ShapeSheet.CellData<bool> LockMoveX { get; set; }
        public VA.ShapeSheet.CellData<bool> LockMoveY { get; set; }
        public VA.ShapeSheet.CellData<bool> LockRotate { get; set; }
        public VA.ShapeSheet.CellData<bool> LockSelect { get; set; }
        public VA.ShapeSheet.CellData<bool> LockTextEdit { get; set; }
        public VA.ShapeSheet.CellData<bool> LockThemeColors { get; set; }
        public VA.ShapeSheet.CellData<bool> LockThemeEffects { get; set; }
        public VA.ShapeSheet.CellData<bool> LockVtxEdit { get; set; }
        public VA.ShapeSheet.CellData<bool> LockWidth { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
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

        private static LockCells get_cells_from_row(LockQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new LockCells();
            cells.LockAspect = qds.GetItem(row, query.LockAspect, v => VA.Convert.DoubleToBool(v));
            cells.LockBegin = qds.GetItem(row, query.LockBegin, v => VA.Convert.DoubleToBool(v));
            cells.LockCalcWH = qds.GetItem(row, query.LockCalcWH, v => VA.Convert.DoubleToBool(v));
            cells.LockCrop = qds.GetItem(row, query.LockCrop, v => VA.Convert.DoubleToBool(v));
            cells.LockCustProp = qds.GetItem(row, query.LockCustProp, v => VA.Convert.DoubleToBool(v));
            cells.LockDelete = qds.GetItem(row, query.LockDelete, v => VA.Convert.DoubleToBool(v));
            cells.LockEnd = qds.GetItem(row, query.LockEnd, v => VA.Convert.DoubleToBool(v));
            cells.LockFormat = qds.GetItem(row, query.LockFormat, v => VA.Convert.DoubleToBool(v));
            cells.LockFromGroupFormat = qds.GetItem(row, query.LockFromGroupFormat, v => VA.Convert.DoubleToBool(v));
            cells.LockGroup = qds.GetItem(row, query.LockGroup, v => VA.Convert.DoubleToBool(v));
            cells.LockHeight = qds.GetItem(row, query.LockHeight, v => VA.Convert.DoubleToBool(v));
            cells.LockMoveX = qds.GetItem(row, query.LockMoveX, v => VA.Convert.DoubleToBool(v));
            cells.LockMoveY = qds.GetItem(row, query.LockMoveY, v => VA.Convert.DoubleToBool(v));
            cells.LockRotate = qds.GetItem(row, query.LockRotate, v => VA.Convert.DoubleToBool(v));
            cells.LockSelect = qds.GetItem(row, query.LockSelect, v => VA.Convert.DoubleToBool(v));
            cells.LockTextEdit = qds.GetItem(row, query.LockTextEdit, v => VA.Convert.DoubleToBool(v));
            cells.LockThemeColors = qds.GetItem(row, query.LockThemeColors, v => VA.Convert.DoubleToBool(v));
            cells.LockThemeEffects = qds.GetItem(row, query.LockThemeEffects, v => VA.Convert.DoubleToBool(v));
            cells.LockVtxEdit = qds.GetItem(row, query.LockVtxEdit, v => VA.Convert.DoubleToBool(v));
            cells.LockWidth = qds.GetItem(row, query.LockWidth, v => VA.Convert.DoubleToBool(v));
            return cells;
        }

        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new LockQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = new LockQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }
    }
}