using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue AvenueSizeX { get; set; }
        public VisioAutomation.Core.CellValue AvenueSizeY { get; set; }
        public VisioAutomation.Core.CellValue BlockSizeX { get; set; }
        public VisioAutomation.Core.CellValue BlockSizeY { get; set; }
        public VisioAutomation.Core.CellValue CtrlAsInput { get; set; }
        public VisioAutomation.Core.CellValue DynamicsOff { get; set; }
        public VisioAutomation.Core.CellValue EnableGrid { get; set; }
        public VisioAutomation.Core.CellValue LineAdjustFrom { get; set; }
        public VisioAutomation.Core.CellValue LineAdjustTo { get; set; }
        public VisioAutomation.Core.CellValue LineJumpCode { get; set; }
        public VisioAutomation.Core.CellValue LineJumpFactorX { get; set; }
        public VisioAutomation.Core.CellValue LineJumpFactorY { get; set; }
        public VisioAutomation.Core.CellValue LineJumpStyle { get; set; }
        public VisioAutomation.Core.CellValue LineRouteExt { get; set; }
        public VisioAutomation.Core.CellValue LineToLineX { get; set; }
        public VisioAutomation.Core.CellValue LineToLineY { get; set; }
        public VisioAutomation.Core.CellValue LineToNodeX { get; set; }
        public VisioAutomation.Core.CellValue LineToNodeY { get; set; }
        public VisioAutomation.Core.CellValue LineJumpDirX { get; set; }
        public VisioAutomation.Core.CellValue LineJumpDirY { get; set; }
        public VisioAutomation.Core.CellValue PageShapeSplit { get; set; }
        public VisioAutomation.Core.CellValue PlaceDepth { get; set; }
        public VisioAutomation.Core.CellValue PlaceFlip { get; set; }
        public VisioAutomation.Core.CellValue PlaceStyle { get; set; }
        public VisioAutomation.Core.CellValue PlowCode { get; set; }
        public VisioAutomation.Core.CellValue ResizePage { get; set; }
        public VisioAutomation.Core.CellValue RouteStyle { get; set; }
        public VisioAutomation.Core.CellValue AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.AvenueSizeX), VisioAutomation.Core.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
            yield return this.Create(nameof(this.AvenueSizeY), VisioAutomation.Core.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
            yield return this.Create(nameof(this.BlockSizeX), VisioAutomation.Core.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
            yield return this.Create(nameof(this.BlockSizeY), VisioAutomation.Core.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
            yield return this.Create(nameof(this.CtrlAsInput), VisioAutomation.Core.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
            yield return this.Create(nameof(this.DynamicsOff), VisioAutomation.Core.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
            yield return this.Create(nameof(this.EnableGrid), VisioAutomation.Core.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
            yield return this.Create(nameof(this.LineAdjustFrom), VisioAutomation.Core.SrcConstants.PageLayoutLineAdjustFrom,
                this.LineAdjustFrom);
            yield return this.Create(nameof(this.LineAdjustTo), VisioAutomation.Core.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
            yield return this.Create(nameof(this.LineJumpCode), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
            yield return this.Create(nameof(this.LineJumpFactorX), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpFactorX,
                this.LineJumpFactorX);
            yield return this.Create(nameof(this.LineJumpFactorY), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpFactorY,
                this.LineJumpFactorY);
            yield return this.Create(nameof(this.LineJumpStyle), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
            yield return this.Create(nameof(this.LineRouteExt), VisioAutomation.Core.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
            yield return this.Create(nameof(this.LineToLineX), VisioAutomation.Core.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
            yield return this.Create(nameof(this.LineToLineY), VisioAutomation.Core.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
            yield return this.Create(nameof(this.LineToNodeX), VisioAutomation.Core.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
            yield return this.Create(nameof(this.LineToNodeY), VisioAutomation.Core.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
            yield return this.Create(nameof(this.LineJumpDirX), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
            yield return this.Create(nameof(this.LineJumpDirY), VisioAutomation.Core.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
            yield return this.Create(nameof(this.PageShapeSplit), VisioAutomation.Core.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
            yield return this.Create(nameof(this.PlaceDepth), VisioAutomation.Core.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            yield return this.Create(nameof(this.PlaceFlip), VisioAutomation.Core.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            yield return this.Create(nameof(this.PlaceStyle), VisioAutomation.Core.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            yield return this.Create(nameof(this.PlowCode), VisioAutomation.Core.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            yield return this.Create(nameof(this.ResizePage), VisioAutomation.Core.SrcConstants.PageLayoutResizePage, this.ResizePage);
            yield return this.Create(nameof(this.RouteStyle), VisioAutomation.Core.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
            yield return this.Create(nameof(this.AvoidPageBreaks), VisioAutomation.Core.SrcConstants.PageLayoutAvoidPageBreaks,
                this.AvoidPageBreaks);
        }

        public static PageLayoutCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = PageLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageLayoutCellsBuilder> PageLayoutCells_lazy_builder = new System.Lazy<PageLayoutCellsBuilder>();


        class PageLayoutCellsBuilder : VASS.CellGroups.CellGroupBuilder<PageLayoutCells>
        {
            public PageLayoutCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }


            public override PageLayoutCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new PageLayoutCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);


                cells.AvenueSizeX = getcellvalue(nameof(PageLayoutCells.AvenueSizeX));
                cells.AvenueSizeY = getcellvalue(nameof(PageLayoutCells.AvenueSizeY));
                cells.BlockSizeX = getcellvalue(nameof(PageLayoutCells.BlockSizeX));
                cells.BlockSizeY = getcellvalue(nameof(PageLayoutCells.BlockSizeY));
                cells.CtrlAsInput = getcellvalue(nameof(PageLayoutCells.CtrlAsInput));
                cells.DynamicsOff = getcellvalue(nameof(PageLayoutCells.DynamicsOff));
                cells.EnableGrid = getcellvalue(nameof(PageLayoutCells.EnableGrid));
                cells.LineAdjustFrom = getcellvalue(nameof(PageLayoutCells.LineAdjustFrom));
                cells.LineAdjustTo = getcellvalue(nameof(PageLayoutCells.LineAdjustTo));
                cells.LineJumpCode = getcellvalue(nameof(PageLayoutCells.LineJumpCode));
                cells.LineJumpFactorX = getcellvalue(nameof(PageLayoutCells.LineJumpFactorX));
                cells.LineJumpFactorY = getcellvalue(nameof(PageLayoutCells.LineJumpFactorY));
                cells.LineJumpStyle = getcellvalue(nameof(PageLayoutCells.LineJumpStyle));
                cells.LineRouteExt = getcellvalue(nameof(PageLayoutCells.LineRouteExt));
                cells.LineToLineX = getcellvalue(nameof(PageLayoutCells.LineToLineX));
                cells.LineToLineY = getcellvalue(nameof(PageLayoutCells.LineToLineY));
                cells.LineToNodeX = getcellvalue(nameof(PageLayoutCells.LineToNodeX));
                cells.LineToNodeY = getcellvalue(nameof(PageLayoutCells.LineToNodeY));
                cells.LineJumpDirX = getcellvalue(nameof(PageLayoutCells.LineJumpDirX));
                cells.LineJumpDirY = getcellvalue(nameof(PageLayoutCells.LineJumpDirY));
                cells.PageShapeSplit = getcellvalue(nameof(PageLayoutCells.PageShapeSplit));
                cells.PlaceDepth = getcellvalue(nameof(PageLayoutCells.PlaceDepth));
                cells.PlaceFlip = getcellvalue(nameof(PageLayoutCells.PlaceFlip));
                cells.PlaceStyle = getcellvalue(nameof(PageLayoutCells.PlaceStyle));
                cells.PlowCode = getcellvalue(nameof(PageLayoutCells.PlowCode));
                cells.ResizePage = getcellvalue(nameof(PageLayoutCells.ResizePage));
                cells.RouteStyle = getcellvalue(nameof(PageLayoutCells.RouteStyle));
                cells.AvoidPageBreaks = getcellvalue(nameof(PageLayoutCells.AvoidPageBreaks));
                return cells;
            }
        }

    }
}