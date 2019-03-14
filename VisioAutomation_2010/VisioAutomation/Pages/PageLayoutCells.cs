using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : VASS.CellGroups.CellGroupBase
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

        public override IEnumerable<VASS.CellGroups.SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutPlowCode, this.PlowCode);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutResizePage, this.ResizePage);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
                yield return VASS.CellGroups.SrcValuePair.Create(VASS.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks);
            }
        }

        public static PageLayoutCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();

        class PageLayoutCellsReader : VASS.CellGroups.CellGroupReader<PageLayoutCells>
        {
            public VASS.Query.CellColumn AvenueSizeX { get; set; }
            public VASS.Query.CellColumn AvenueSizeY { get; set; }
            public VASS.Query.CellColumn BlockSizeX { get; set; }
            public VASS.Query.CellColumn BlockSizeY { get; set; }
            public VASS.Query.CellColumn ControlAsInput { get; set; }
            public VASS.Query.CellColumn DynamicsOff { get; set; }
            public VASS.Query.CellColumn EnableGrid { get; set; }
            public VASS.Query.CellColumn LineAdjustFrom { get; set; }
            public VASS.Query.CellColumn LineAdjustTo { get; set; }
            public VASS.Query.CellColumn LineJumpCode { get; set; }
            public VASS.Query.CellColumn LineJumpFactorX { get; set; }
            public VASS.Query.CellColumn LineJumpFactorY { get; set; }
            public VASS.Query.CellColumn LineJumpStyle { get; set; }
            public VASS.Query.CellColumn LineRouteExt { get; set; }
            public VASS.Query.CellColumn LineToLineX { get; set; }
            public VASS.Query.CellColumn LineToLineY { get; set; }
            public VASS.Query.CellColumn LineToNodeX { get; set; }
            public VASS.Query.CellColumn LineToNodeY { get; set; }
            public VASS.Query.CellColumn LineJumpDirX { get; set; }
            public VASS.Query.CellColumn LineJumpDirY { get; set; }
            public VASS.Query.CellColumn ShapeSplit { get; set; }
            public VASS.Query.CellColumn PlaceDepth { get; set; }
            public VASS.Query.CellColumn PlaceFlip { get; set; }
            public VASS.Query.CellColumn PlaceStyle { get; set; }
            public VASS.Query.CellColumn PlowCode { get; set; }
            public VASS.Query.CellColumn ResizePage { get; set; }
            public VASS.Query.CellColumn RouteStyle { get; set; }
            public VASS.Query.CellColumn AvoidPageBreaks { get; set; }

            public PageLayoutCellsReader()
            {
                this.AvenueSizeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvenueSizeX, nameof(this.AvenueSizeX));
                this.AvenueSizeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvenueSizeY, nameof(this.AvenueSizeY));
                this.BlockSizeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutBlockSizeX, nameof(this.BlockSizeX));
                this.BlockSizeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutBlockSizeY, nameof(this.BlockSizeY));
                this.ControlAsInput = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutControlAsInput, nameof(this.ControlAsInput));
                this.DynamicsOff = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutDynamicsOff, nameof(this.DynamicsOff));
                this.EnableGrid = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutEnableGrid, nameof(this.EnableGrid));
                this.LineAdjustFrom = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineAdjustFrom, nameof(this.LineAdjustFrom));
                this.LineAdjustTo = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineAdjustTo, nameof(this.LineAdjustTo));
                this.LineJumpCode = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpCode, nameof(this.LineJumpCode));
                this.LineJumpFactorX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpFactorX, nameof(this.LineJumpFactorX));
                this.LineJumpFactorY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpFactorY, nameof(this.LineJumpFactorY));
                this.LineJumpStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpStyle, nameof(this.LineJumpStyle));
                this.LineRouteExt = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineRouteExt, nameof(this.LineRouteExt));
                this.LineToLineX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToLineX, nameof(this.LineToLineX));
                this.LineToLineY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToLineY, nameof(this.LineToLineY));
                this.LineToNodeX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToNodeX, nameof(this.LineToNodeX));
                this.LineToNodeY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineToNodeY, nameof(this.LineToNodeY));
                this.LineJumpDirX = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpDirX, nameof(this.LineJumpDirX));
                this.LineJumpDirY = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutLineJumpDirY, nameof(this.LineJumpDirY));
                this.ShapeSplit = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutShapeSplit, nameof(this.ShapeSplit));
                this.PlaceDepth = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceDepth, nameof(this.PlaceDepth));
                this.PlaceFlip = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceFlip, nameof(this.PlaceFlip));
                this.PlaceStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlaceStyle, nameof(this.PlaceStyle));
                this.PlowCode = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutPlowCode, nameof(this.PlowCode));
                this.ResizePage = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutResizePage, nameof(this.ResizePage));
                this.RouteStyle = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutRouteStyle, nameof(this.RouteStyle));
                this.AvoidPageBreaks = this.query_singlerow.Columns.Add(VASS.SrcConstants.PageLayoutAvoidPageBreaks, nameof(this.AvoidPageBreaks));
            }


            public override PageLayoutCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new PageLayoutCells();
                cells.AvenueSizeX = row[this.AvenueSizeX];
                cells.AvenueSizeY = row[this.AvenueSizeY];
                cells.BlockSizeX = row[this.BlockSizeX];
                cells.BlockSizeY = row[this.BlockSizeY];
                cells.CtrlAsInput = row[this.ControlAsInput];
                cells.DynamicsOff = row[this.DynamicsOff];
                cells.EnableGrid = row[this.EnableGrid];
                cells.LineAdjustFrom = row[this.LineAdjustFrom];
                cells.LineAdjustTo = row[this.LineAdjustTo];
                cells.LineJumpCode = row[this.LineJumpCode];
                cells.LineJumpFactorX = row[this.LineJumpFactorX];
                cells.LineJumpFactorY = row[this.LineJumpFactorY];
                cells.LineJumpStyle = row[this.LineJumpStyle];
                cells.LineRouteExt = row[this.LineRouteExt];
                cells.LineToLineX = row[this.LineToLineX];
                cells.LineToLineY = row[this.LineToLineY];
                cells.LineToNodeX = row[this.LineToNodeX];
                cells.LineToNodeY = row[this.LineToNodeY];
                cells.LineJumpDirX = row[this.LineJumpDirX];
                cells.LineJumpDirY = row[this.LineJumpDirY];
                cells.PageShapeSplit = row[this.ShapeSplit];
                cells.PlaceDepth = row[this.PlaceDepth];
                cells.PlaceFlip = row[this.PlaceFlip];
                cells.PlaceStyle = row[this.PlaceStyle];
                cells.PlowCode = row[this.PlowCode];
                cells.ResizePage = row[this.ResizePage];
                cells.RouteStyle = row[this.RouteStyle];
                cells.AvoidPageBreaks = row[this.AvoidPageBreaks];
                return cells;
            }
        }

    }
}