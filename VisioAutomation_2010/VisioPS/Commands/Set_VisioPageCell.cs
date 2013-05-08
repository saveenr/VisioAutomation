using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageCell")]
    public class Set_VisioPageCell: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public string PageWidth { get; set; }
        
        [SMA.Parameter(Mandatory = false)] 
        public string PageHeight { get; set; }
        
        [SMA.Parameter(Mandatory = false)] 
        public string PageBottomMargin;
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLeftMargin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLineJumpDirX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLineJumpDirY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageRightMargin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageScale { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageShapeSplit { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageTopMargin { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CenterX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CenterY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string PaperKind { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PrintGrid { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PrintPageOrientation { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ScaleX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ScaleY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PaperSource { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScaleType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScale { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingSizeType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string InhibitSnap { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwObliqueAngle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwScaleFactor { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string UIVisibility { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridSpacing { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XRulerDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XRulerOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridSpacing { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YRulerDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YRulerOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string CtrlAsInput { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DynamicsOff { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string EnableGrid { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustFrom { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustTo { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpCode { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineRouteExt { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceDepth { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceFlip { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ResizePage { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlowCode { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string RouteStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvoidPageBreaks { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingResizeType { get; set; }
 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }
        
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

                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, this.PageHeight);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageScale, this.PageScale);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, this.PageWidth);

                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.CenterX, this.CenterX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.CenterY, this.CenterY);

                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PaperKind, this.PaperKind);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ScaleX, this.ScaleX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ScaleY, this.ScaleY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PaperSource, this.PaperSource);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ShdwType, this.ShdwType);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.PlowCode, this.PlowCode);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.ResizePage, this.ResizePage);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks);
                update.SetFormulaIgnoreNull(VisioAutomation.ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType);


                this.WriteVerboseEx("BlastGuards: {0}", this.BlastGuards);
                this.WriteVerboseEx("TestCircular: {0}", this.TestCircular);
                this.WriteVerboseEx("Number of Shapes : {0}", 1);
                this.WriteVerboseEx("Number of Total Updates: {0}", update.Count());
                this.WriteVerboseEx("Number of Updates per Shape: {0}", update.Count() / 1);

                using (var undoscope = new VA.Application.UndoScope(this.ScriptingSession.VisioApplication, "SetPageCells"))
                {
                    this.WriteVerboseEx("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerboseEx("End Update");
                }
            }

        }
    }
}