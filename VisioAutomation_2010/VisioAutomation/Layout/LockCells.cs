using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
{
    public class LockCells : VA.ShapeSheet.CellGroups.CellGroupEx
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

        public override void ApplyFormulas(ApplyFormula func)
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


        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(shape, query, query.GetCells);
        }


        private static LockQuery m_query;
        private static LockQuery get_query()
        {
            m_query = m_query ?? new LockQuery();
            return m_query;
        }

        class LockQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int LockAspect { get; set; }
            public int LockBegin { get; set; }
            public int LockCalcWH { get; set; }
            public int LockCrop { get; set; }
            public int LockCustProp { get; set; }
            public int LockDelete { get; set; }
            public int LockEnd { get; set; }
            public int LockFormat { get; set; }
            public int LockFromGroupFormat { get; set; }
            public int LockGroup { get; set; }
            public int LockHeight { get; set; }
            public int LockMoveX { get; set; }
            public int LockMoveY { get; set; }
            public int LockRotate { get; set; }
            public int LockSelect { get; set; }
            public int LockTextEdit { get; set; }
            public int LockThemeColors { get; set; }
            public int LockThemeEffects { get; set; }
            public int LockVtxEdit { get; set; }
            public int LockWidth { get; set; }

            public LockQuery() :
                base()
            {
                this.LockAspect = this.AddCell(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect");
                this.LockBegin = this.AddCell(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin");
                this.LockCalcWH = this.AddCell(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
                this.LockCrop = this.AddCell(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop");
                this.LockCustProp = this.AddCell(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
                this.LockDelete = this.AddCell(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete");
                this.LockEnd = this.AddCell(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd");
                this.LockFormat = this.AddCell(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat");
                this.LockFromGroupFormat = this.AddCell(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
                this.LockGroup = this.AddCell(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup");
                this.LockHeight = this.AddCell(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight");
                this.LockMoveX = this.AddCell(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
                this.LockMoveY = this.AddCell(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
                this.LockRotate = this.AddCell(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate");
                this.LockSelect = this.AddCell(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect");
                this.LockTextEdit = this.AddCell(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
                this.LockThemeColors = this.AddCell(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
                this.LockThemeEffects = this.AddCell(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
                this.LockVtxEdit = this.AddCell(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
                this.LockWidth = this.AddCell(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth");
            }

            public LockCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;
                var cells = new LockCells();
                cells.LockAspect = row[this.LockAspect].ToBool();
                cells.LockBegin = row[this.LockBegin].ToBool();
                cells.LockCalcWH = row[this.LockCalcWH].ToBool();
                cells.LockCrop = row[this.LockCrop].ToBool();
                cells.LockCustProp = row[this.LockCustProp].ToBool();
                cells.LockDelete = row[this.LockDelete].ToBool();
                cells.LockEnd = row[this.LockEnd].ToBool();
                cells.LockFormat = row[this.LockFormat].ToBool();
                cells.LockFromGroupFormat = row[this.LockFromGroupFormat].ToBool();
                cells.LockGroup = row[this.LockGroup].ToBool();
                cells.LockHeight = row[this.LockHeight].ToBool();
                cells.LockMoveX = row[this.LockMoveX].ToBool();
                cells.LockMoveY = row[this.LockMoveY].ToBool();
                cells.LockRotate = row[this.LockRotate].ToBool();
                cells.LockSelect = row[this.LockSelect].ToBool();
                cells.LockTextEdit = row[this.LockTextEdit].ToBool();
                cells.LockThemeColors = row[this.LockThemeColors].ToBool();
                cells.LockThemeEffects = row[this.LockThemeEffects].ToBool();
                cells.LockVtxEdit = row[this.LockVtxEdit].ToBool();
                cells.LockWidth = row[this.LockWidth].ToBool();
                return cells;
            }
        }
    }
}