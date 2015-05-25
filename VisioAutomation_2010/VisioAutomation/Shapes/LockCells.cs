using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VAQUERY=VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : ShapeSheet.CellGroups.CellGroup
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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
            }
        }


        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = LockCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = LockCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<LockCellQuery> lazy_query = new System.Lazy<LockCellQuery>();


        class LockCellQuery : VAQUERY.CellQuery
        {
            public VAQUERY.CellColumn LockAspect { get; set; }
            public VAQUERY.CellColumn LockBegin { get; set; }
            public VAQUERY.CellColumn LockCalcWH { get; set; }
            public VAQUERY.CellColumn LockCrop { get; set; }
            public VAQUERY.CellColumn LockCustProp { get; set; }
            public VAQUERY.CellColumn LockDelete { get; set; }
            public VAQUERY.CellColumn LockEnd { get; set; }
            public VAQUERY.CellColumn LockFormat { get; set; }
            public VAQUERY.CellColumn LockFromGroupFormat { get; set; }
            public VAQUERY.CellColumn LockGroup { get; set; }
            public VAQUERY.CellColumn LockHeight { get; set; }
            public VAQUERY.CellColumn LockMoveX { get; set; }
            public VAQUERY.CellColumn LockMoveY { get; set; }
            public VAQUERY.CellColumn LockRotate { get; set; }
            public VAQUERY.CellColumn LockSelect { get; set; }
            public VAQUERY.CellColumn LockTextEdit { get; set; }
            public VAQUERY.CellColumn LockThemeColors { get; set; }
            public VAQUERY.CellColumn LockThemeEffects { get; set; }
            public VAQUERY.CellColumn LockVtxEdit { get; set; }
            public VAQUERY.CellColumn LockWidth { get; set; }

            public LockCellQuery() 
            {
                this.LockAspect = this.AddCell(ShapeSheet.SRCConstants.LockAspect, "LockAspect");
                this.LockBegin = this.AddCell(ShapeSheet.SRCConstants.LockBegin, "LockBegin");
                this.LockCalcWH = this.AddCell(ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
                this.LockCrop = this.AddCell(ShapeSheet.SRCConstants.LockCrop, "LockCrop");
                this.LockCustProp = this.AddCell(ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
                this.LockDelete = this.AddCell(ShapeSheet.SRCConstants.LockDelete, "LockDelete");
                this.LockEnd = this.AddCell(ShapeSheet.SRCConstants.LockEnd, "LockEnd");
                this.LockFormat = this.AddCell(ShapeSheet.SRCConstants.LockFormat, "LockFormat");
                this.LockFromGroupFormat = this.AddCell(ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
                this.LockGroup = this.AddCell(ShapeSheet.SRCConstants.LockGroup, "LockGroup");
                this.LockHeight = this.AddCell(ShapeSheet.SRCConstants.LockHeight, "LockHeight");
                this.LockMoveX = this.AddCell(ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
                this.LockMoveY = this.AddCell(ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
                this.LockRotate = this.AddCell(ShapeSheet.SRCConstants.LockRotate, "LockRotate");
                this.LockSelect = this.AddCell(ShapeSheet.SRCConstants.LockSelect, "LockSelect");
                this.LockTextEdit = this.AddCell(ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
                this.LockThemeColors = this.AddCell(ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
                this.LockThemeEffects = this.AddCell(ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
                this.LockVtxEdit = this.AddCell(ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
                this.LockWidth = this.AddCell(ShapeSheet.SRCConstants.LockWidth, "LockWidth");

            }

            public LockCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {
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