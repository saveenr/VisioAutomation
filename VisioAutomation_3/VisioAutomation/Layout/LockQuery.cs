using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
{
    class LockQuery : VA.ShapeSheet.Query.CellQuery
    {
        public VA.ShapeSheet.Query.CellQueryColumn LockAspect { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockBegin { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockCalcWH { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockCrop { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockCustProp { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockDelete { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockEnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockFormat { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockFromGroupFormat { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockGroup { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockHeight { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockMoveX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockMoveY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockRotate { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockSelect { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockTextEdit { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockThemeColors { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockThemeEffects { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockVtxEdit { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LockWidth { get; set; }

        public LockQuery() :
            base()
        {
            this.LockAspect = this.AddColumn(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect");
            this.LockBegin = this.AddColumn(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin");
            this.LockCalcWH = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
            this.LockCrop = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop");
            this.LockCustProp = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
            this.LockDelete = this.AddColumn(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete");
            this.LockEnd = this.AddColumn(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd");
            this.LockFormat = this.AddColumn(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat");
            this.LockFromGroupFormat = this.AddColumn(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
            this.LockGroup = this.AddColumn(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup");
            this.LockHeight = this.AddColumn(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight");
            this.LockMoveX = this.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
            this.LockMoveY = this.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
            this.LockRotate = this.AddColumn(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate");
            this.LockSelect = this.AddColumn(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect");
            this.LockTextEdit = this.AddColumn(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
            this.LockThemeColors = this.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
            this.LockThemeEffects = this.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
            this.LockVtxEdit = this.AddColumn(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
            this.LockWidth = this.AddColumn(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth");
        }
    }
}