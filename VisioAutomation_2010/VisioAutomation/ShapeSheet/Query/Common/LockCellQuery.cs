namespace VisioAutomation.ShapeSheet.Query.Common
{
    class LockCellQuery : CellQuery
    {
        public Query.CellColumn LockAspect { get; set; }
        public Query.CellColumn LockBegin { get; set; }
        public Query.CellColumn LockCalcWH { get; set; }
        public Query.CellColumn LockCrop { get; set; }
        public Query.CellColumn LockCustProp { get; set; }
        public Query.CellColumn LockDelete { get; set; }
        public Query.CellColumn LockEnd { get; set; }
        public Query.CellColumn LockFormat { get; set; }
        public Query.CellColumn LockFromGroupFormat { get; set; }
        public Query.CellColumn LockGroup { get; set; }
        public Query.CellColumn LockHeight { get; set; }
        public Query.CellColumn LockMoveX { get; set; }
        public Query.CellColumn LockMoveY { get; set; }
        public Query.CellColumn LockRotate { get; set; }
        public Query.CellColumn LockSelect { get; set; }
        public Query.CellColumn LockTextEdit { get; set; }
        public Query.CellColumn LockThemeColors { get; set; }
        public Query.CellColumn LockThemeEffects { get; set; }
        public Query.CellColumn LockVtxEdit { get; set; }
        public Query.CellColumn LockWidth { get; set; }

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

        public VisioAutomation.Shapes.LockCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.LockCells();
            cells.LockAspect = Extensions.CellDataMethods.ToBool(row[this.LockAspect]);
            cells.LockBegin = Extensions.CellDataMethods.ToBool(row[this.LockBegin]);
            cells.LockCalcWH = Extensions.CellDataMethods.ToBool(row[this.LockCalcWH]);
            cells.LockCrop = Extensions.CellDataMethods.ToBool(row[this.LockCrop]);
            cells.LockCustProp = Extensions.CellDataMethods.ToBool(row[this.LockCustProp]);
            cells.LockDelete = Extensions.CellDataMethods.ToBool(row[this.LockDelete]);
            cells.LockEnd = Extensions.CellDataMethods.ToBool(row[this.LockEnd]);
            cells.LockFormat = Extensions.CellDataMethods.ToBool(row[this.LockFormat]);
            cells.LockFromGroupFormat = Extensions.CellDataMethods.ToBool(row[this.LockFromGroupFormat]);
            cells.LockGroup = Extensions.CellDataMethods.ToBool(row[this.LockGroup]);
            cells.LockHeight = Extensions.CellDataMethods.ToBool(row[this.LockHeight]);
            cells.LockMoveX = Extensions.CellDataMethods.ToBool(row[this.LockMoveX]);
            cells.LockMoveY = Extensions.CellDataMethods.ToBool(row[this.LockMoveY]);
            cells.LockRotate = Extensions.CellDataMethods.ToBool(row[this.LockRotate]);
            cells.LockSelect = Extensions.CellDataMethods.ToBool(row[this.LockSelect]);
            cells.LockTextEdit = Extensions.CellDataMethods.ToBool(row[this.LockTextEdit]);
            cells.LockThemeColors = Extensions.CellDataMethods.ToBool(row[this.LockThemeColors]);
            cells.LockThemeEffects = Extensions.CellDataMethods.ToBool(row[this.LockThemeEffects]);
            cells.LockVtxEdit = Extensions.CellDataMethods.ToBool(row[this.LockVtxEdit]);
            cells.LockWidth = Extensions.CellDataMethods.ToBool(row[this.LockWidth]);
            return cells;
        }
    }
}