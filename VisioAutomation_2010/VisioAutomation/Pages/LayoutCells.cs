using System.Collections.Generic;
using VACG=VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class LayoutCells : VACG.CellGroup
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

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.AvenueSizeX), Core.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
            yield return this._create(nameof(this.AvenueSizeY), Core.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
            yield return this._create(nameof(this.BlockSizeX), Core.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
            yield return this._create(nameof(this.BlockSizeY), Core.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
            yield return this._create(nameof(this.CtrlAsInput), Core.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
            yield return this._create(nameof(this.DynamicsOff), Core.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
            yield return this._create(nameof(this.EnableGrid), Core.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
            yield return this._create(nameof(this.LineAdjustFrom), Core.SrcConstants.PageLayoutLineAdjustFrom,
                this.LineAdjustFrom);
            yield return this._create(nameof(this.LineAdjustTo), Core.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
            yield return this._create(nameof(this.LineJumpCode), Core.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
            yield return this._create(nameof(this.LineJumpFactorX), Core.SrcConstants.PageLayoutLineJumpFactorX,
                this.LineJumpFactorX);
            yield return this._create(nameof(this.LineJumpFactorY), Core.SrcConstants.PageLayoutLineJumpFactorY,
                this.LineJumpFactorY);
            yield return this._create(nameof(this.LineJumpStyle), Core.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
            yield return this._create(nameof(this.LineRouteExt), Core.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
            yield return this._create(nameof(this.LineToLineX), Core.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
            yield return this._create(nameof(this.LineToLineY), Core.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
            yield return this._create(nameof(this.LineToNodeX), Core.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
            yield return this._create(nameof(this.LineToNodeY), Core.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
            yield return this._create(nameof(this.LineJumpDirX), Core.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
            yield return this._create(nameof(this.LineJumpDirY), Core.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
            yield return this._create(nameof(this.PageShapeSplit), Core.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
            yield return this._create(nameof(this.PlaceDepth), Core.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            yield return this._create(nameof(this.PlaceFlip), Core.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            yield return this._create(nameof(this.PlaceStyle), Core.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            yield return this._create(nameof(this.PlowCode), Core.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            yield return this._create(nameof(this.ResizePage), Core.SrcConstants.PageLayoutResizePage, this.ResizePage);
            yield return this._create(nameof(this.RouteStyle), Core.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
            yield return this._create(nameof(this.AvoidPageBreaks), Core.SrcConstants.PageLayoutAvoidPageBreaks,
                this.AvoidPageBreaks);
        }

        public static LayoutCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();


        class Builder : VACG.CellGroupBuilder<LayoutCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }


            public override LayoutCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new LayoutCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);


                cells.AvenueSizeX = getcellvalue(nameof(AvenueSizeX));
                cells.AvenueSizeY = getcellvalue(nameof(AvenueSizeY));
                cells.BlockSizeX = getcellvalue(nameof(BlockSizeX));
                cells.BlockSizeY = getcellvalue(nameof(BlockSizeY));
                cells.CtrlAsInput = getcellvalue(nameof(CtrlAsInput));
                cells.DynamicsOff = getcellvalue(nameof(DynamicsOff));
                cells.EnableGrid = getcellvalue(nameof(EnableGrid));
                cells.LineAdjustFrom = getcellvalue(nameof(LineAdjustFrom));
                cells.LineAdjustTo = getcellvalue(nameof(LineAdjustTo));
                cells.LineJumpCode = getcellvalue(nameof(LineJumpCode));
                cells.LineJumpFactorX = getcellvalue(nameof(LineJumpFactorX));
                cells.LineJumpFactorY = getcellvalue(nameof(LineJumpFactorY));
                cells.LineJumpStyle = getcellvalue(nameof(LineJumpStyle));
                cells.LineRouteExt = getcellvalue(nameof(LineRouteExt));
                cells.LineToLineX = getcellvalue(nameof(LineToLineX));
                cells.LineToLineY = getcellvalue(nameof(LineToLineY));
                cells.LineToNodeX = getcellvalue(nameof(LineToNodeX));
                cells.LineToNodeY = getcellvalue(nameof(LineToNodeY));
                cells.LineJumpDirX = getcellvalue(nameof(LineJumpDirX));
                cells.LineJumpDirY = getcellvalue(nameof(LineJumpDirY));
                cells.PageShapeSplit = getcellvalue(nameof(PageShapeSplit));
                cells.PlaceDepth = getcellvalue(nameof(PlaceDepth));
                cells.PlaceFlip = getcellvalue(nameof(PlaceFlip));
                cells.PlaceStyle = getcellvalue(nameof(PlaceStyle));
                cells.PlowCode = getcellvalue(nameof(PlowCode));
                cells.ResizePage = getcellvalue(nameof(ResizePage));
                cells.RouteStyle = getcellvalue(nameof(RouteStyle));
                cells.AvoidPageBreaks = getcellvalue(nameof(AvoidPageBreaks));
                return cells;
            }
        }

    }
}