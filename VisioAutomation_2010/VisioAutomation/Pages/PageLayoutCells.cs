using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;

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

                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.AvenueSizeX), VASS.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.AvenueSizeY), VASS.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.BlockSizeX), VASS.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.BlockSizeY), VASS.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.CtrlAsInput), VASS.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.DynamicsOff), VASS.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.EnableGrid), VASS.SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineAdjustFrom), VASS.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineAdjustTo), VASS.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpCode), VASS.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpFactorX), VASS.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpFactorY), VASS.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpStyle), VASS.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineRouteExt), VASS.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineToLineX), VASS.SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineToLineY), VASS.SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineToNodeX), VASS.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineToNodeY), VASS.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpDirX), VASS.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LineJumpDirY), VASS.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PageShapeSplit), VASS.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PlaceDepth), VASS.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PlaceFlip), VASS.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PlaceStyle), VASS.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.PlowCode), VASS.SrcConstants.PageLayoutPlowCode, this.PlowCode);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.ResizePage), VASS.SrcConstants.PageLayoutResizePage, this.ResizePage);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.RouteStyle), VASS.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.AvoidPageBreaks), VASS.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks);

            }
        }
    }
}