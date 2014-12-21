using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageCell")]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false,Position=0)] 
        public System.Collections.Hashtable Hashtable  { get; set; }

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
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new[] { this.client.Page.Get() };

            var valuemap = new CellValueMap(Get_VisioPageCell.GetPageCellDictionary());

            valuemap.UpdateValueMap(this.Hashtable);

            valuemap.SetIf("PageBottomMargin",this.PageBottomMargin);
            valuemap.SetIf("PageHeight",this.PageHeight);
            valuemap.SetIf("PageLeftMargin",this.PageLeftMargin);
            valuemap.SetIf("PageLineJumpDirX",this.PageLineJumpDirX);
            valuemap.SetIf("PageLineJumpDirY",this.PageLineJumpDirY);
            valuemap.SetIf("PageRightMargin",this.PageRightMargin);
            valuemap.SetIf("PageScale",this.PageScale);
            valuemap.SetIf("PageShapeSplit",this.PageShapeSplit);
            valuemap.SetIf("PageTopMargin",this.PageTopMargin);
            valuemap.SetIf("PageWidth",this.PageWidth);
            valuemap.SetIf("CenterX",this.CenterX);
            valuemap.SetIf("CenterY",this.CenterY);
            valuemap.SetIf("PaperKind",this.PaperKind);
            valuemap.SetIf("PrintGrid",this.PrintGrid);
            valuemap.SetIf("PrintPageOrientation",this.PrintPageOrientation);
            valuemap.SetIf("ScaleX",this.ScaleX);
            valuemap.SetIf("ScaleY",this.ScaleY);
            valuemap.SetIf("PaperSource",this.PaperSource);
            valuemap.SetIf("DrawingScale",this.DrawingScale);
            valuemap.SetIf("DrawingScaleType",this.DrawingScaleType);
            valuemap.SetIf("DrawingSizeType",this.DrawingSizeType);
            valuemap.SetIf("InhibitSnap",this.InhibitSnap);
            valuemap.SetIf("ShdwObliqueAngle",this.ShdwObliqueAngle);
            valuemap.SetIf("ShdwOffsetX",this.ShdwOffsetX);
            valuemap.SetIf("ShdwOffsetY",this.ShdwOffsetY);
            valuemap.SetIf("ShdwScaleFactor",this.ShdwScaleFactor);
            valuemap.SetIf("ShdwType",this.ShdwType);
            valuemap.SetIf("UIVisibility",this.UIVisibility);
            valuemap.SetIf("XGridDensity",this.XGridDensity);
            valuemap.SetIf("XGridOrigin",this.XGridOrigin);
            valuemap.SetIf("XGridSpacing",this.XGridSpacing);
            valuemap.SetIf("XRulerDensity",this.XRulerDensity);
            valuemap.SetIf("XRulerOrigin",this.XRulerOrigin);
            valuemap.SetIf("YGridDensity",this.YGridDensity);
            valuemap.SetIf("YGridOrigin",this.YGridOrigin);
            valuemap.SetIf("YGridSpacing",this.YGridSpacing);
            valuemap.SetIf("YRulerDensity",this.YRulerDensity);
            valuemap.SetIf("YRulerOrigin",this.YRulerOrigin);
            valuemap.SetIf("AvenueSizeX",this.AvenueSizeX);
            valuemap.SetIf("AvenueSizeY",this.AvenueSizeY);
            valuemap.SetIf("BlockSizeX",this.BlockSizeX);
            valuemap.SetIf("BlockSizeY",this.BlockSizeY);
            valuemap.SetIf("CtrlAsInput",this.CtrlAsInput);
            valuemap.SetIf("DynamicsOff",this.DynamicsOff);
            valuemap.SetIf("EnableGrid",this.EnableGrid);
            valuemap.SetIf("LineAdjustFrom",this.LineAdjustFrom);
            valuemap.SetIf("LineAdjustTo",this.LineAdjustTo);
            valuemap.SetIf("LineJumpCode",this.LineJumpCode);
            valuemap.SetIf("LineJumpFactorX",this.LineJumpFactorX);
            valuemap.SetIf("LineJumpFactorY",this.LineJumpFactorY);
            valuemap.SetIf("LineJumpStyle",this.LineJumpStyle);
            valuemap.SetIf("LineRouteExt",this.LineRouteExt);
            valuemap.SetIf("LineToLineX",this.LineToLineX);
            valuemap.SetIf("LineToLineY",this.LineToLineY);
            valuemap.SetIf("LineToNodeX",this.LineToNodeX);
            valuemap.SetIf("LineToNodeY",this.LineToNodeY);
            valuemap.SetIf("PageLineJumpDirX",this.PageLineJumpDirX);
            valuemap.SetIf("PageLineJumpDirY",this.PageLineJumpDirY);
            valuemap.SetIf("PageShapeSplit",this.PageShapeSplit);
            valuemap.SetIf("PlaceDepth",this.PlaceDepth);
            valuemap.SetIf("PlaceFlip",this.PlaceFlip);
            valuemap.SetIf("PlaceStyle",this.PlaceStyle);
            valuemap.SetIf("PlowCode",this.PlowCode);
            valuemap.SetIf("ResizePage",this.ResizePage);
            valuemap.SetIf("RouteStyle",this.RouteStyle);
            valuemap.SetIf("AvoidPageBreaks",this.AvoidPageBreaks);
            valuemap.SetIf("DrawingResizeType",this.DrawingResizeType);


            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormulaIgnoreNull( cell_src , cell_value);
                }
                this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
                this.WriteVerbose("TestCircular: {0}", this.TestCircular);
                this.WriteVerbose("Number of Shapes : {0}", 1);
                this.WriteVerbose("Number of Total Updates: {0}", update.Count());
                this.WriteVerbose("Number of Updates per Shape: {0}", update.Count() / 1);

                using (var undoscope = new VA.Application.UndoScope(this.client.VisioApplication, "SetPageCells"))
                {
                    this.WriteVerbose("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerbose("End Update");
                }
            }

        }

    }
}