using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPageCell")]
    public class Get_VisioPageCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public string[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter AvenueSizeX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter AvenueSizeY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter AvoidPageBreaks { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlockSizeX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlockSizeY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CenterX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CenterY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CtrlAsInput { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DrawingResizeType { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DrawingScale { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DrawingScaleType { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DrawingSizeType { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DynamicsOff { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter EnableGrid { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter InhibitSnap { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineAdjustFrom { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineAdjustTo { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineJumpCode { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineJumpFactorX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineJumpFactorY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineJumpStyle { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineRouteExt { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineToLineX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineToLineY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineToNodeX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineToNodeY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageBottomMargin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageHeight { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageLeftMargin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageLineJumpDirX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageLineJumpDirY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageRightMargin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageScale { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageShapeSplit { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageTopMargin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PageWidth { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PaperKind { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PaperSource { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PlaceDepth { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PlaceFlip { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PlaceStyle { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PlowCode { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PrintGrid { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PrintPageOrientation { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ResizePage { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter RouteStyle { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ScaleX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ScaleY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwObliqueAngle { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwOffsetX { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwOffsetY { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwScaleFactor { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwType { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter UIVisibility { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter XGridDensity { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter XGridOrigin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter XGridSpacing { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter XRulerDensity { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter XRulerOrigin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter YGridDensity { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter YGridOrigin { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter YGridSpacing { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter YRulerDensity { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter YRulerOrigin { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var query = new VA.ShapeSheet.Query.CellQuery();

            addcell(query, this.AvenueSizeX, "AvenueSizeX");
            addcell(query, this.AvenueSizeY, "AvenueSizeY");
            addcell(query, this.AvoidPageBreaks, "AvoidPageBreaks");
            addcell(query, this.BlockSizeX, "BlockSizeX");
            addcell(query, this.BlockSizeY, "BlockSizeY");
            addcell(query, this.CenterX, "CenterX");
            addcell(query, this.CenterY, "CenterY");
            addcell(query, this.CtrlAsInput, "CtrlAsInput");
            addcell(query, this.DrawingResizeType, "DrawingResizeType");
            addcell(query, this.DrawingScale, "DrawingScale");
            addcell(query, this.DrawingScaleType, "DrawingScaleType");
            addcell(query, this.DrawingSizeType, "DrawingSizeType");
            addcell(query, this.DynamicsOff, "DynamicsOff");
            addcell(query, this.EnableGrid, "EnableGrid");
            addcell(query, this.InhibitSnap, "InhibitSnap");
            addcell(query, this.LineAdjustFrom, "LineAdjustFrom");
            addcell(query, this.LineAdjustTo, "LineAdjustTo");
            addcell(query, this.LineJumpCode, "LineJumpCode");
            addcell(query, this.LineJumpFactorX, "LineJumpFactorX");
            addcell(query, this.LineJumpFactorY, "LineJumpFactorY");
            addcell(query, this.LineJumpStyle, "LineJumpStyle");
            addcell(query, this.LineRouteExt, "LineRouteExt");
            addcell(query, this.LineToLineX, "LineToLineX");
            addcell(query, this.LineToLineY, "LineToLineY");
            addcell(query, this.LineToNodeX, "LineToNodeX");
            addcell(query, this.LineToNodeY, "LineToNodeY");
            addcell(query, this.PageBottomMargin, "PageBottomMargin");
            addcell(query, this.PageHeight, "PageHeight");
            addcell(query, this.PageLeftMargin, "PageLeftMargin");
            addcell(query, this.PageLineJumpDirX, "PageLineJumpDirX");
            addcell(query, this.PageLineJumpDirY, "PageLineJumpDirY");
            addcell(query, this.PageRightMargin, "PageRightMargin");
            addcell(query, this.PageScale, "PageScale");
            addcell(query, this.PageShapeSplit, "PageShapeSplit");
            addcell(query, this.PageTopMargin, "PageTopMargin");
            addcell(query, this.PageWidth, "PageWidth");
            addcell(query, this.PaperKind, "PaperKind");
            addcell(query, this.PaperSource, "PaperSource");
            addcell(query, this.PlaceDepth, "PlaceDepth");
            addcell(query, this.PlaceFlip, "PlaceFlip");
            addcell(query, this.PlaceStyle, "PlaceStyle");
            addcell(query, this.PlowCode, "PlowCode");
            addcell(query, this.PrintGrid, "PrintGrid");
            addcell(query, this.PrintPageOrientation, "PrintPageOrientation");
            addcell(query, this.ResizePage, "ResizePage");
            addcell(query, this.RouteStyle, "RouteStyle");
            addcell(query, this.ScaleX, "ScaleX");
            addcell(query, this.ScaleY, "ScaleY");
            addcell(query, this.ShdwObliqueAngle, "ShdwObliqueAngle");
            addcell(query, this.ShdwOffsetX, "ShdwOffsetX");
            addcell(query, this.ShdwOffsetY, "ShdwOffsetY");
            addcell(query, this.ShdwScaleFactor, "ShdwScaleFactor");
            addcell(query, this.ShdwType, "ShdwType");
            addcell(query, this.UIVisibility, "UIVisibility");
            addcell(query, this.XGridDensity, "XGridDensity");
            addcell(query, this.XGridOrigin, "XGridOrigin");
            addcell(query, this.XGridSpacing, "XGridSpacing");
            addcell(query, this.XRulerDensity, "XRulerDensity");
            addcell(query, this.XRulerOrigin, "XRulerOrigin");
            addcell(query, this.YGridDensity, "YGridDensity");
            addcell(query, this.YGridOrigin, "YGridOrigin");
            addcell(query, this.YGridSpacing, "YGridSpacing");
            addcell(query, this.YRulerDensity, "YRulerDensity");
            addcell(query, this.YRulerOrigin, "YRulerOrigin");

            var dic = GetPageCellDictionary();
            SetFromCellNames(query, this.Cells, dic);

            var surface = new VA.Drawing.DrawingSurface(this.client.Page.Get());
            
            var target_shapeids = new[] { surface.Page.ID };

            this.WriteVerbose("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerbose("Start Query");

            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);

            this.WriteObject(dt);
            this.WriteVerbose("End Query");
        }

        public static void SetFromCellNames(VA.ShapeSheet.Query.CellQuery query, string[] Cells, CellMap dic)
        {
            if (Cells == null)
            {
                return;
            }

            foreach (string resolved_cellname in dic.ResolveNames(Cells))
            {
                if (!query.Columns.Contains(resolved_cellname))
                {
                    query.Columns.Add(dic[resolved_cellname], resolved_cellname);
                }
            }
        }

        private void addcell(VA.ShapeSheet.Query.CellQuery query, bool switchpar, string cellname)
        {
            var dic = Get_VisioPageCell.GetPageCellDictionary();
            if (switchpar)
            {
                query.Columns.Add(dic[cellname], cellname);
            }
        }

        private static CellMap cellmap;


        public static CellMap GetPageCellDictionary()
        {
            if (cellmap == null)
            {
                cellmap = new CellMap();
                cellmap["PageBottomMargin"] = VA.ShapeSheet.SRCConstants.PageBottomMargin;
                cellmap["PageHeight"] = VA.ShapeSheet.SRCConstants.PageHeight;
                cellmap["PageLeftMargin"] = VA.ShapeSheet.SRCConstants.PageLeftMargin;
                cellmap["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;
                cellmap["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;
                cellmap["PageRightMargin"] = VA.ShapeSheet.SRCConstants.PageRightMargin;
                cellmap["PageScale"] = VA.ShapeSheet.SRCConstants.PageScale;
                cellmap["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;
                cellmap["PageTopMargin"] = VA.ShapeSheet.SRCConstants.PageTopMargin;
                cellmap["PageWidth"] = VA.ShapeSheet.SRCConstants.PageWidth;
                cellmap["CenterX"] = VA.ShapeSheet.SRCConstants.CenterX;
                cellmap["CenterY"] = VA.ShapeSheet.SRCConstants.CenterY;
                cellmap["PaperKind"] = VA.ShapeSheet.SRCConstants.PaperKind;
                cellmap["PrintGrid"] = VA.ShapeSheet.SRCConstants.PrintGrid;
                cellmap["PrintPageOrientation"] = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
                cellmap["ScaleX"] = VA.ShapeSheet.SRCConstants.ScaleX;
                cellmap["ScaleY"] = VA.ShapeSheet.SRCConstants.ScaleY;
                cellmap["PaperSource"] = VA.ShapeSheet.SRCConstants.PaperSource;
                cellmap["DrawingScale"] = VA.ShapeSheet.SRCConstants.DrawingScale;
                cellmap["DrawingScaleType"] = VA.ShapeSheet.SRCConstants.DrawingScaleType;
                cellmap["DrawingSizeType"] = VA.ShapeSheet.SRCConstants.DrawingSizeType;
                cellmap["InhibitSnap"] = VA.ShapeSheet.SRCConstants.InhibitSnap;
                cellmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                cellmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                cellmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                cellmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                cellmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                cellmap["UIVisibility"] = VA.ShapeSheet.SRCConstants.UIVisibility;
                cellmap["XGridDensity"] = VA.ShapeSheet.SRCConstants.XGridDensity;
                cellmap["XGridOrigin"] = VA.ShapeSheet.SRCConstants.XGridOrigin;
                cellmap["XGridSpacing"] = VA.ShapeSheet.SRCConstants.XGridSpacing;
                cellmap["XRulerDensity"] = VA.ShapeSheet.SRCConstants.XRulerDensity;
                cellmap["XRulerOrigin"] = VA.ShapeSheet.SRCConstants.XRulerOrigin;
                cellmap["YGridDensity"] = VA.ShapeSheet.SRCConstants.YGridDensity;
                cellmap["YGridOrigin"] = VA.ShapeSheet.SRCConstants.YGridOrigin;
                cellmap["YGridSpacing"] = VA.ShapeSheet.SRCConstants.YGridSpacing;
                cellmap["YRulerDensity"] = VA.ShapeSheet.SRCConstants.YRulerDensity;
                cellmap["YRulerOrigin"] = VA.ShapeSheet.SRCConstants.YRulerOrigin;
                cellmap["AvenueSizeX"] = VA.ShapeSheet.SRCConstants.AvenueSizeX;
                cellmap["AvenueSizeY"] = VA.ShapeSheet.SRCConstants.AvenueSizeY;
                cellmap["BlockSizeX"] = VA.ShapeSheet.SRCConstants.BlockSizeX;
                cellmap["BlockSizeY"] = VA.ShapeSheet.SRCConstants.BlockSizeY;
                cellmap["CtrlAsInput"] = VA.ShapeSheet.SRCConstants.CtrlAsInput;
                cellmap["DynamicsOff"] = VA.ShapeSheet.SRCConstants.DynamicsOff;
                cellmap["EnableGrid"] = VA.ShapeSheet.SRCConstants.EnableGrid;
                cellmap["LineAdjustFrom"] = VA.ShapeSheet.SRCConstants.LineAdjustFrom;
                cellmap["LineAdjustTo"] = VA.ShapeSheet.SRCConstants.LineAdjustTo;
                cellmap["LineJumpCode"] = VA.ShapeSheet.SRCConstants.LineJumpCode;
                cellmap["LineJumpFactorX"] = VA.ShapeSheet.SRCConstants.LineJumpFactorX;
                cellmap["LineJumpFactorY"] = VA.ShapeSheet.SRCConstants.LineJumpFactorY;
                cellmap["LineJumpStyle"] = VA.ShapeSheet.SRCConstants.LineJumpStyle;
                cellmap["LineRouteExt"] = VA.ShapeSheet.SRCConstants.LineRouteExt;
                cellmap["LineToLineX"] = VA.ShapeSheet.SRCConstants.LineToLineX;
                cellmap["LineToLineY"] = VA.ShapeSheet.SRCConstants.LineToLineY;
                cellmap["LineToNodeX"] = VA.ShapeSheet.SRCConstants.LineToNodeX;
                cellmap["LineToNodeY"] = VA.ShapeSheet.SRCConstants.LineToNodeY;
                cellmap["PlaceDepth"] = VA.ShapeSheet.SRCConstants.PlaceDepth;
                cellmap["PlaceFlip"] = VA.ShapeSheet.SRCConstants.PlaceFlip;
                cellmap["PlaceStyle"] = VA.ShapeSheet.SRCConstants.PlaceStyle;
                cellmap["PlowCode"] = VA.ShapeSheet.SRCConstants.PlowCode;
                cellmap["ResizePage"] = VA.ShapeSheet.SRCConstants.ResizePage;
                cellmap["RouteStyle"] = VA.ShapeSheet.SRCConstants.RouteStyle;
                //cellmap["AvoidPageBreaks"] = VA.ShapeSheet.SRCConstants.AvoidPageBreaks;
                //cellmap["DrawingResizeType"] = VA.ShapeSheet.SRCConstants.DrawingResizeType;
            }
            return cellmap;
        }

        /*
       
AvenueSizeX
AvenueSizeY
AvoidPageBreaks
BlockSizeX
BlockSizeY
CenterX
CenterY
CtrlAsInput
DrawingResizeType
DrawingScale
DrawingScaleType
DrawingSizeType
DynamicsOff
EnableGrid
InhibitSnap
LineAdjustFrom
LineAdjustTo
LineJumpCode
LineJumpFactorX
LineJumpFactorY
LineJumpStyle
LineRouteExt
LineToLineX
LineToLineY
LineToNodeX
LineToNodeY
PageBottomMargin
PageHeight
PageLeftMargin
PageLineJumpDirX
PageLineJumpDirY
PageRightMargin
PageScale
PageShapeSplit
PageTopMargin
PageWidth
PaperKind
PaperSource
PlaceDepth
PlaceFlip
PlaceStyle
PlowCode
PrintGrid
PrintPageOrientation
ResizePage
RouteStyle
ScaleX
ScaleY
ShdwObliqueAngle
ShdwOffsetX
ShdwOffsetY
ShdwScaleFactor
ShdwType
UIVisibility
XGridDensity
XGridOrigin
XGridSpacing
XRulerDensity
XRulerOrigin
YGridDensity
YGridOrigin
YGridSpacing
YRulerDensity
YRulerOrigin
 
         
         * */
    }
}