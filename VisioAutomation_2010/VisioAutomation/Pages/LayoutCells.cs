using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class LayoutCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.AvenueSizeX), Core.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
            yield return this.Create(nameof(this.AvenueSizeY), Core.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
            yield return this.Create(nameof(this.BlockSizeX), Core.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
            yield return this.Create(nameof(this.BlockSizeY), Core.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
            yield return this.Create(nameof(this.CtrlAsInput), Core.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
            yield return this.Create(nameof(this.DynamicsOff), Core.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
            yield return this.Create(nameof(this.EnableGrid), Core.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
            yield return this.Create(nameof(this.LineAdjustFrom), Core.SrcConstants.PageLayoutLineAdjustFrom,
                this.LineAdjustFrom);
            yield return this.Create(nameof(this.LineAdjustTo), Core.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
            yield return this.Create(nameof(this.LineJumpCode), Core.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
            yield return this.Create(nameof(this.LineJumpFactorX), Core.SrcConstants.PageLayoutLineJumpFactorX,
                this.LineJumpFactorX);
            yield return this.Create(nameof(this.LineJumpFactorY), Core.SrcConstants.PageLayoutLineJumpFactorY,
                this.LineJumpFactorY);
            yield return this.Create(nameof(this.LineJumpStyle), Core.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
            yield return this.Create(nameof(this.LineRouteExt), Core.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
            yield return this.Create(nameof(this.LineToLineX), Core.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
            yield return this.Create(nameof(this.LineToLineY), Core.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
            yield return this.Create(nameof(this.LineToNodeX), Core.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
            yield return this.Create(nameof(this.LineToNodeY), Core.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
            yield return this.Create(nameof(this.LineJumpDirX), Core.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
            yield return this.Create(nameof(this.LineJumpDirY), Core.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
            yield return this.Create(nameof(this.PageShapeSplit), Core.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
            yield return this.Create(nameof(this.PlaceDepth), Core.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            yield return this.Create(nameof(this.PlaceFlip), Core.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            yield return this.Create(nameof(this.PlaceStyle), Core.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            yield return this.Create(nameof(this.PlowCode), Core.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            yield return this.Create(nameof(this.ResizePage), Core.SrcConstants.PageLayoutResizePage, this.ResizePage);
            yield return this.Create(nameof(this.RouteStyle), Core.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
            yield return this.Create(nameof(this.AvoidPageBreaks), Core.SrcConstants.PageLayoutAvoidPageBreaks,
                this.AvoidPageBreaks);
        }

        public static LayoutCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = PageLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageLayoutCellsBuilder> PageLayoutCells_lazy_builder = new System.Lazy<PageLayoutCellsBuilder>();


        class PageLayoutCellsBuilder : CellGroupBuilder<LayoutCells>
        {
            public PageLayoutCellsBuilder() : base(CellGroupBuilderType.SingleRow)
            {
            }


            public override LayoutCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new LayoutCells();
                var getcellvalue = row_to_cellgroup(row, cols);


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