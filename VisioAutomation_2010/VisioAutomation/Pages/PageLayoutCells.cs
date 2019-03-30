using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral AvenueSizeX { get; set; }
        public VASS.CellValueLiteral AvenueSizeY { get; set; }
        public VASS.CellValueLiteral BlockSizeX { get; set; }
        public VASS.CellValueLiteral BlockSizeY { get; set; }
        public VASS.CellValueLiteral CtrlAsInput { get; set; }
        public VASS.CellValueLiteral DynamicsOff { get; set; }
        public VASS.CellValueLiteral EnableGrid { get; set; }
        public VASS.CellValueLiteral LineAdjustFrom { get; set; }
        public VASS.CellValueLiteral LineAdjustTo { get; set; }
        public VASS.CellValueLiteral LineJumpCode { get; set; }
        public VASS.CellValueLiteral LineJumpFactorX { get; set; }
        public VASS.CellValueLiteral LineJumpFactorY { get; set; }
        public VASS.CellValueLiteral LineJumpStyle { get; set; }
        public VASS.CellValueLiteral LineRouteExt { get; set; }
        public VASS.CellValueLiteral LineToLineX { get; set; }
        public VASS.CellValueLiteral LineToLineY { get; set; }
        public VASS.CellValueLiteral LineToNodeX { get; set; }
        public VASS.CellValueLiteral LineToNodeY { get; set; }
        public VASS.CellValueLiteral LineJumpDirX { get; set; }
        public VASS.CellValueLiteral LineJumpDirY { get; set; }
        public VASS.CellValueLiteral PageShapeSplit { get; set; }
        public VASS.CellValueLiteral PlaceDepth { get; set; }
        public VASS.CellValueLiteral PlaceFlip { get; set; }
        public VASS.CellValueLiteral PlaceStyle { get; set; }
        public VASS.CellValueLiteral PlowCode { get; set; }
        public VASS.CellValueLiteral ResizePage { get; set; }
        public VASS.CellValueLiteral RouteStyle { get; set; }
        public VASS.CellValueLiteral AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {

                yield return this.Create(nameof(this.AvenueSizeX), VASS.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
                yield return this.Create(nameof(this.AvenueSizeY), VASS.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
                yield return this.Create(nameof(this.BlockSizeX), VASS.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
                yield return this.Create(nameof(this.BlockSizeY), VASS.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
                yield return this.Create(nameof(this.CtrlAsInput), VASS.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
                yield return this.Create(nameof(this.DynamicsOff), VASS.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
                yield return this.Create(nameof(this.EnableGrid), VASS.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
                yield return this.Create(nameof(this.LineAdjustFrom), VASS.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom);
                yield return this.Create(nameof(this.LineAdjustTo), VASS.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
                yield return this.Create(nameof(this.LineJumpCode), VASS.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
                yield return this.Create(nameof(this.LineJumpFactorX), VASS.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX);
                yield return this.Create(nameof(this.LineJumpFactorY), VASS.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY);
                yield return this.Create(nameof(this.LineJumpStyle), VASS.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
                yield return this.Create(nameof(this.LineRouteExt), VASS.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
                yield return this.Create(nameof(this.LineToLineX), VASS.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
                yield return this.Create(nameof(this.LineToLineY), VASS.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
                yield return this.Create(nameof(this.LineToNodeX), VASS.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
                yield return this.Create(nameof(this.LineToNodeY), VASS.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
                yield return this.Create(nameof(this.LineJumpDirX), VASS.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
                yield return this.Create(nameof(this.LineJumpDirY), VASS.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
                yield return this.Create(nameof(this.PageShapeSplit), VASS.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
                yield return this.Create(nameof(this.PlaceDepth), VASS.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
                yield return this.Create(nameof(this.PlaceFlip), VASS.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
                yield return this.Create(nameof(this.PlaceStyle), VASS.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
                yield return this.Create(nameof(this.PlowCode), VASS.SrcConstants.PageLayoutPlowCode, this.PlowCode);
                yield return this.Create(nameof(this.ResizePage), VASS.SrcConstants.PageLayoutResizePage, this.ResizePage);
                yield return this.Create(nameof(this.RouteStyle), VASS.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
                yield return this.Create(nameof(this.AvoidPageBreaks), VASS.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks);

            }
        }

        public static PageLayoutCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
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