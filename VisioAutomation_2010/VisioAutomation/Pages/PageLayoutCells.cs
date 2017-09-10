using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
namespace VisioAutomation.Pages
{
    public class PageLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral AvenueSizeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral AvenueSizeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BlockSizeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BlockSizeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CtrlAsInput { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DynamicsOff { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral EnableGrid { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineAdjustFrom { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineAdjustTo { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpFactorX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpFactorY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineRouteExt { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToLineX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToLineY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToNodeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToNodeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PageShapeSplit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceDepth { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceFlip { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlowCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ResizePage { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RouteStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks);
            }
        }

        public static PageLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, CellValueType cvt)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();

        class PageLayoutCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageLayoutCells>
        {
            public CellColumn AvenueSizeX { get; set; }
            public CellColumn AvenueSizeY { get; set; }
            public CellColumn BlockSizeX { get; set; }
            public CellColumn BlockSizeY { get; set; }
            public CellColumn ControlAsInput { get; set; }
            public CellColumn DynamicsOff { get; set; }
            public CellColumn EnableGrid { get; set; }
            public CellColumn LineAdjustFrom { get; set; }
            public CellColumn LineAdjustTo { get; set; }
            public CellColumn LineJumpCode { get; set; }
            public CellColumn LineJumpFactorX { get; set; }
            public CellColumn LineJumpFactorY { get; set; }
            public CellColumn LineJumpStyle { get; set; }
            public CellColumn LineRouteExt { get; set; }
            public CellColumn LineToLineX { get; set; }
            public CellColumn LineToLineY { get; set; }
            public CellColumn LineToNodeX { get; set; }
            public CellColumn LineToNodeY { get; set; }
            public CellColumn LineJumpDirX { get; set; }
            public CellColumn LineJumpDirY { get; set; }
            public CellColumn ShapeSplit { get; set; }
            public CellColumn PlaceDepth { get; set; }
            public CellColumn PlaceFlip { get; set; }
            public CellColumn PlaceStyle { get; set; }
            public CellColumn PlowCode { get; set; }
            public CellColumn ResizePage { get; set; }
            public CellColumn RouteStyle { get; set; }
            public CellColumn AvoidPageBreaks { get; set; }

            public PageLayoutCellsReader()
            {
                this.AvenueSizeX = this.query.Columns.Add(SrcConstants.PageLayoutAvenueSizeX, nameof(SrcConstants.PageLayoutAvenueSizeX));
                this.AvenueSizeY = this.query.Columns.Add(SrcConstants.PageLayoutAvenueSizeY, nameof(SrcConstants.PageLayoutAvenueSizeY));
                this.BlockSizeX = this.query.Columns.Add(SrcConstants.PageLayoutBlockSizeX, nameof(SrcConstants.PageLayoutBlockSizeX));
                this.BlockSizeY = this.query.Columns.Add(SrcConstants.PageLayoutBlockSizeY, nameof(SrcConstants.PageLayoutBlockSizeY));
                this.ControlAsInput = this.query.Columns.Add(SrcConstants.PageLayoutControlAsInput, nameof(SrcConstants.PageLayoutControlAsInput));
                this.DynamicsOff = this.query.Columns.Add(SrcConstants.PageLayoutDynamicsOff, nameof(SrcConstants.PageLayoutDynamicsOff));
                this.EnableGrid = this.query.Columns.Add(SrcConstants.PageLayoutEnableGrid, nameof(SrcConstants.PageLayoutEnableGrid));
                this.LineAdjustFrom = this.query.Columns.Add(SrcConstants.PageLayoutLineAdjustFrom, nameof(SrcConstants.PageLayoutLineAdjustFrom));
                this.LineAdjustTo = this.query.Columns.Add(SrcConstants.PageLayoutLineAdjustTo, nameof(SrcConstants.PageLayoutLineAdjustTo));
                this.LineJumpCode = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpCode, nameof(SrcConstants.PageLayoutLineJumpCode));
                this.LineJumpFactorX = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpFactorX, nameof(SrcConstants.PageLayoutLineJumpFactorX));
                this.LineJumpFactorY = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpFactorY, nameof(SrcConstants.PageLayoutLineJumpFactorY));
                this.LineJumpStyle = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpStyle, nameof(SrcConstants.PageLayoutLineJumpStyle));
                this.LineRouteExt = this.query.Columns.Add(SrcConstants.PageLayoutLineRouteExt, nameof(SrcConstants.PageLayoutLineRouteExt));
                this.LineToLineX = this.query.Columns.Add(SrcConstants.PageLayoutLineToLineX, nameof(SrcConstants.PageLayoutLineToLineX));
                this.LineToLineY = this.query.Columns.Add(SrcConstants.PageLayoutLineToLineY, nameof(SrcConstants.PageLayoutLineToLineY));
                this.LineToNodeX = this.query.Columns.Add(SrcConstants.PageLayoutLineToNodeX, nameof(SrcConstants.PageLayoutLineToNodeX));
                this.LineToNodeY = this.query.Columns.Add(SrcConstants.PageLayoutLineToNodeY, nameof(SrcConstants.PageLayoutLineToNodeY));
                this.LineJumpDirX = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpDirX, nameof(SrcConstants.PageLayoutLineJumpDirX));
                this.LineJumpDirY = this.query.Columns.Add(SrcConstants.PageLayoutLineJumpDirY, nameof(SrcConstants.PageLayoutLineJumpDirY));
                this.ShapeSplit = this.query.Columns.Add(SrcConstants.PageLayoutShapeSplit, nameof(SrcConstants.PageLayoutShapeSplit));
                this.PlaceDepth = this.query.Columns.Add(SrcConstants.PageLayoutPlaceDepth, nameof(SrcConstants.PageLayoutPlaceDepth));
                this.PlaceFlip = this.query.Columns.Add(SrcConstants.PageLayoutPlaceFlip, nameof(SrcConstants.PageLayoutPlaceFlip));
                this.PlaceStyle = this.query.Columns.Add(SrcConstants.PageLayoutPlaceStyle, nameof(SrcConstants.PageLayoutPlaceStyle));
                this.PlowCode = this.query.Columns.Add(SrcConstants.PageLayoutPlowCode, nameof(SrcConstants.PageLayoutPlowCode));
                this.ResizePage = this.query.Columns.Add(SrcConstants.PageLayoutResizePage, nameof(SrcConstants.PageLayoutResizePage));
                this.RouteStyle = this.query.Columns.Add(SrcConstants.PageLayoutRouteStyle, nameof(SrcConstants.PageLayoutRouteStyle));
                this.AvoidPageBreaks = this.query.Columns.Add(SrcConstants.PageLayoutAvoidPageBreaks, nameof(SrcConstants.PageLayoutAvoidPageBreaks));
            }


            public override Pages.PageLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Pages.PageLayoutCells();
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