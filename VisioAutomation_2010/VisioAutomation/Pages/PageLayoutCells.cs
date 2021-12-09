using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : CellRecord
    {
        public Core.CellValue AvenueSizeX { get; set; }
        public Core.CellValue AvenueSizeY { get; set; }
        public Core.CellValue BlockSizeX { get; set; }
        public Core.CellValue BlockSizeY { get; set; }
        public Core.CellValue CtrlAsInput { get; set; }
        public Core.CellValue DynamicsOff { get; set; }
        public Core.CellValue EnableGrid { get; set; }
        public Core.CellValue LineAdjustFrom { get; set; }
        public Core.CellValue LineAdjustTo { get; set; }
        public Core.CellValue LineJumpCode { get; set; }
        public Core.CellValue LineJumpFactorX { get; set; }
        public Core.CellValue LineJumpFactorY { get; set; }
        public Core.CellValue LineJumpStyle { get; set; }
        public Core.CellValue LineRouteExt { get; set; }
        public Core.CellValue LineToLineX { get; set; }
        public Core.CellValue LineToLineY { get; set; }
        public Core.CellValue LineToNodeX { get; set; }
        public Core.CellValue LineToNodeY { get; set; }
        public Core.CellValue LineJumpDirX { get; set; }
        public Core.CellValue LineJumpDirY { get; set; }
        public Core.CellValue PageShapeSplit { get; set; }
        public Core.CellValue PlaceDepth { get; set; }
        public Core.CellValue PlaceFlip { get; set; }
        public Core.CellValue PlaceStyle { get; set; }
        public Core.CellValue PlowCode { get; set; }
        public Core.CellValue ResizePage { get; set; }
        public Core.CellValue RouteStyle { get; set; }
        public Core.CellValue AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.AvenueSizeX), Core.SrcConstants.PageLayoutAvenueSizeX,
                this.AvenueSizeX);
            yield return this._create(nameof(this.AvenueSizeY), Core.SrcConstants.PageLayoutAvenueSizeY,
                this.AvenueSizeY);
            yield return this._create(nameof(this.BlockSizeX), Core.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
            yield return this._create(nameof(this.BlockSizeY), Core.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
            yield return this._create(nameof(this.CtrlAsInput), Core.SrcConstants.PageLayoutControlAsInput,
                this.CtrlAsInput);
            yield return this._create(nameof(this.DynamicsOff), Core.SrcConstants.PageLayoutDynamicsOff,
                this.DynamicsOff);
            yield return this._create(nameof(this.EnableGrid), Core.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
            yield return this._create(nameof(this.LineAdjustFrom), Core.SrcConstants.PageLayoutLineAdjustFrom,
                this.LineAdjustFrom);
            yield return this._create(nameof(this.LineAdjustTo), Core.SrcConstants.PageLayoutLineAdjustTo,
                this.LineAdjustTo);
            yield return this._create(nameof(this.LineJumpCode), Core.SrcConstants.PageLayoutLineJumpCode,
                this.LineJumpCode);
            yield return this._create(nameof(this.LineJumpFactorX), Core.SrcConstants.PageLayoutLineJumpFactorX,
                this.LineJumpFactorX);
            yield return this._create(nameof(this.LineJumpFactorY), Core.SrcConstants.PageLayoutLineJumpFactorY,
                this.LineJumpFactorY);
            yield return this._create(nameof(this.LineJumpStyle), Core.SrcConstants.PageLayoutLineJumpStyle,
                this.LineJumpStyle);
            yield return this._create(nameof(this.LineRouteExt), Core.SrcConstants.PageLayoutLineRouteExt,
                this.LineRouteExt);
            yield return this._create(nameof(this.LineToLineX), Core.SrcConstants.PageLayoutLineToLineX,
                this.LineToLineX);
            yield return this._create(nameof(this.LineToLineY), Core.SrcConstants.PageLayoutLineToLineY,
                this.LineToLineY);
            yield return this._create(nameof(this.LineToNodeX), Core.SrcConstants.PageLayoutLineToNodeX,
                this.LineToNodeX);
            yield return this._create(nameof(this.LineToNodeY), Core.SrcConstants.PageLayoutLineToNodeY,
                this.LineToNodeY);
            yield return this._create(nameof(this.LineJumpDirX), Core.SrcConstants.PageLayoutLineJumpDirX,
                this.LineJumpDirX);
            yield return this._create(nameof(this.LineJumpDirY), Core.SrcConstants.PageLayoutLineJumpDirY,
                this.LineJumpDirY);
            yield return this._create(nameof(this.PageShapeSplit), Core.SrcConstants.PageLayoutShapeSplit,
                this.PageShapeSplit);
            yield return this._create(nameof(this.PlaceDepth), Core.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            yield return this._create(nameof(this.PlaceFlip), Core.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            yield return this._create(nameof(this.PlaceStyle), Core.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            yield return this._create(nameof(this.PlowCode), Core.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            yield return this._create(nameof(this.ResizePage), Core.SrcConstants.PageLayoutResizePage, this.ResizePage);
            yield return this._create(nameof(this.RouteStyle), Core.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
            yield return this._create(nameof(this.AvoidPageBreaks), Core.SrcConstants.PageLayoutAvoidPageBreaks,
                this.AvoidPageBreaks);
        }

        public static PageLayoutCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static PageLayoutCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new PageLayoutCells();
            var getcellvalue = getvalfromrowfunc(row, cols);


            record.AvenueSizeX = getcellvalue(nameof(AvenueSizeX));
            record.AvenueSizeY = getcellvalue(nameof(AvenueSizeY));
            record.BlockSizeX = getcellvalue(nameof(BlockSizeX));
            record.BlockSizeY = getcellvalue(nameof(BlockSizeY));
            record.CtrlAsInput = getcellvalue(nameof(CtrlAsInput));
            record.DynamicsOff = getcellvalue(nameof(DynamicsOff));
            record.EnableGrid = getcellvalue(nameof(EnableGrid));
            record.LineAdjustFrom = getcellvalue(nameof(LineAdjustFrom));
            record.LineAdjustTo = getcellvalue(nameof(LineAdjustTo));
            record.LineJumpCode = getcellvalue(nameof(LineJumpCode));
            record.LineJumpFactorX = getcellvalue(nameof(LineJumpFactorX));
            record.LineJumpFactorY = getcellvalue(nameof(LineJumpFactorY));
            record.LineJumpStyle = getcellvalue(nameof(LineJumpStyle));
            record.LineRouteExt = getcellvalue(nameof(LineRouteExt));
            record.LineToLineX = getcellvalue(nameof(LineToLineX));
            record.LineToLineY = getcellvalue(nameof(LineToLineY));
            record.LineToNodeX = getcellvalue(nameof(LineToNodeX));
            record.LineToNodeY = getcellvalue(nameof(LineToNodeY));
            record.LineJumpDirX = getcellvalue(nameof(LineJumpDirX));
            record.LineJumpDirY = getcellvalue(nameof(LineJumpDirY));
            record.PageShapeSplit = getcellvalue(nameof(PageShapeSplit));
            record.PlaceDepth = getcellvalue(nameof(PlaceDepth));
            record.PlaceFlip = getcellvalue(nameof(PlaceFlip));
            record.PlaceStyle = getcellvalue(nameof(PlaceStyle));
            record.PlowCode = getcellvalue(nameof(PlowCode));
            record.ResizePage = getcellvalue(nameof(ResizePage));
            record.RouteStyle = getcellvalue(nameof(RouteStyle));
            record.AvoidPageBreaks = getcellvalue(nameof(AvoidPageBreaks));
            return record;
        }
        class Builder : CellRecordBuilderCellQuery<PageLayoutCells>
        {
            public Builder() : base(PageLayoutCells.RowToRecord)
            {
            }
        }
    }
}