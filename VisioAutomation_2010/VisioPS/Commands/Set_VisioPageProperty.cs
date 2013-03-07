using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageProperty")]
    public class Set_VisioPageProperty: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public string Width { get; set; }
        [SMA.Parameter(Mandatory = false)] public string Height { get; set; }
        [SMA.Parameter(Mandatory = false)] public string PageBottomMargin;
        [SMA.Parameter(Mandatory = false)] public string PageHeight;
        [SMA.Parameter(Mandatory = false)] public string PageLeftMargin;
        [SMA.Parameter(Mandatory = false)] public string PageLineJumpDirX;
        [SMA.Parameter(Mandatory = false)] public string PageLineJumpDirY;
        [SMA.Parameter(Mandatory = false)] public string PageRightMargin;
        [SMA.Parameter(Mandatory = false)] public string PageScale;
        [SMA.Parameter(Mandatory = false)] public string PageShapeSplit;
        [SMA.Parameter(Mandatory = false)] public string PageTopMargin;

        [SMA.Parameter(Mandatory = false)]
        public string CenterX;

        [SMA.Parameter(Mandatory = false)]
        public string CenterY;
        [SMA.Parameter(Mandatory = false)] public string PageWidth;
        [SMA.Parameter(Mandatory = false)]
        public string PaperKind;
        [SMA.Parameter(Mandatory = false)]
        public string PrintGrid;
        [SMA.Parameter(Mandatory = false)]
        public string PrintPageOrientation;
        [SMA.Parameter(Mandatory = false)]
        public string ScaleX;
        [SMA.Parameter(Mandatory = false)]
        public string ScaleY;
        [SMA.Parameter(Mandatory = false)]
        public string PaperSource;
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScaleType;
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScale;
        [SMA.Parameter(Mandatory = false)]
        public string DrawingSizeType;
        [SMA.Parameter(Mandatory = false)]
        public string InhibitSnap;
        [SMA.Parameter(Mandatory = false)]
        public string ShdwObliqueAngle;
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetX;
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetY;
        [SMA.Parameter(Mandatory = false)]
        public string ShdwScaleFactor;
        [SMA.Parameter(Mandatory = false)]
        public string ShdwType;
        [SMA.Parameter(Mandatory = false)]
        public string UIVisibility;
        [SMA.Parameter(Mandatory = false)]
        public string XGridDensity;
        [SMA.Parameter(Mandatory = false)]
        public string XGridOrigin;
        [SMA.Parameter(Mandatory = false)]
        public string XGridSpacing;
        [SMA.Parameter(Mandatory = false)]
        public string XRulerDensity;
        [SMA.Parameter(Mandatory = false)]
        public string XRulerOrigin;
        [SMA.Parameter(Mandatory = false)]
        public string YGridDensity;
        [SMA.Parameter(Mandatory = false)]
        public string YGridOrigin;
        [SMA.Parameter(Mandatory = false)]
        public string YGridSpacing;
        [SMA.Parameter(Mandatory = false)]
        public string YRulerDensity;
        [SMA.Parameter(Mandatory = false)]
        public string YRulerOrigin;
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeX;
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeY;
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeX;
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeY;
        [SMA.Parameter(Mandatory = false)]
        public string CtrlAsInput;
        [SMA.Parameter(Mandatory = false)]
        public string DynamicsOff;
        [SMA.Parameter(Mandatory = false)]
        public string EnableGrid;
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustFrom;
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustTo;
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpCode;
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorX;
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorY;
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpStyle;
        [SMA.Parameter(Mandatory = false)]
        public string LineRouteExt;
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineX;
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineY;
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeX;
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeY;
        [SMA.Parameter(Mandatory = false)]
        public string PlaceDepth;
        [SMA.Parameter(Mandatory = false)]
        public string PlaceFlip;
        [SMA.Parameter(Mandatory = false)]
        public string PlaceStyle;
        [SMA.Parameter(Mandatory = false)]
        public string ResizePage;
        [SMA.Parameter(Mandatory = false)]
        public string PlowCode;
        [SMA.Parameter(Mandatory = false)]
        public string RouteStyle;
        [SMA.Parameter(Mandatory = false)]
        public string AvoidPageBreaks;
        [SMA.Parameter(Mandatory = false)]
        public string DrawingResizeType;
 

        [SMA.Parameter(Mandatory = false)] public IList<IVisio.Page> Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;


        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new [] { scriptingsession.Page.Get() };

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;
                var id = pagesheet.ID16;
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageHeight, this.Height);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageScale, this.PageScale);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageWidth, this.PageWidth);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CenterX, this.CenterX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CenterY, this.CenterY);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PaperKind, this.PaperKind);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ScaleX, this.ScaleX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ScaleY, this.ScaleY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PaperSource, this.PaperSource);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwType, this.ShdwType);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PlowCode, this.PlowCode);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ResizePage, this.ResizePage);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType);

                update.Execute(page);
            }

        }
    }
}