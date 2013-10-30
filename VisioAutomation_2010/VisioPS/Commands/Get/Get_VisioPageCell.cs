using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPageCell")]
    public class Get_VisioPageCell: VisioPSCmdlet
    {

        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	AvenueSizeX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	AvenueSizeY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	AvoidPageBreaks	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	BlockSizeX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	BlockSizeY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	CenterX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	CenterY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	CtrlAsInput	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	DrawingResizeType	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	DrawingScale	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	DrawingScaleType	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	DrawingSizeType	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	DynamicsOff	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	EnableGrid	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	InhibitSnap	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineAdjustFrom	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineAdjustTo	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineJumpCode	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineJumpFactorX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineJumpFactorY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineJumpStyle	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineRouteExt	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineToLineX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineToLineY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineToNodeX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	LineToNodeY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageBottomMargin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageHeight	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageLeftMargin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageLineJumpDirX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageLineJumpDirY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageRightMargin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageScale	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageShapeSplit	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageTopMargin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PageWidth	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PaperKind	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PaperSource	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PlaceDepth	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PlaceFlip	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PlaceStyle	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PlowCode	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PrintGrid	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	PrintPageOrientation	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ResizePage	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	RouteStyle	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ScaleX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ScaleY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ShdwObliqueAngle	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ShdwOffsetX	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ShdwOffsetY	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ShdwScaleFactor	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	ShdwType	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	UIVisibility	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	XGridDensity	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	XGridOrigin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	XGridSpacing	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	XRulerDensity	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	XRulerOrigin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	YGridDensity	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	YGridOrigin	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	YGridSpacing	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	YRulerDensity	{ get; set; }	
[SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter	YRulerOrigin	{ get; set; }	




        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var query = new VA.ShapeSheet.Query.CellQuery();

    
addcell(query,this.AvenueSizeX,"AvenueSizeX");
addcell(query,this.AvenueSizeY,"AvenueSizeY");
addcell(query,this.AvoidPageBreaks,"AvoidPageBreaks");
addcell(query,this.BlockSizeX,"BlockSizeX");
addcell(query,this.BlockSizeY,"BlockSizeY");
addcell(query,this.CenterX,"CenterX");
addcell(query,this.CenterY,"CenterY");
addcell(query,this.CtrlAsInput,"CtrlAsInput");
addcell(query,this.DrawingResizeType,"DrawingResizeType");
addcell(query,this.DrawingScale,"DrawingScale");
addcell(query,this.DrawingScaleType,"DrawingScaleType");
addcell(query,this.DrawingSizeType,"DrawingSizeType");
addcell(query,this.DynamicsOff,"DynamicsOff");
addcell(query,this.EnableGrid,"EnableGrid");
addcell(query,this.InhibitSnap,"InhibitSnap");
addcell(query,this.LineAdjustFrom,"LineAdjustFrom");
addcell(query,this.LineAdjustTo,"LineAdjustTo");
addcell(query,this.LineJumpCode,"LineJumpCode");
addcell(query,this.LineJumpFactorX,"LineJumpFactorX");
addcell(query,this.LineJumpFactorY,"LineJumpFactorY");
addcell(query,this.LineJumpStyle,"LineJumpStyle");
addcell(query,this.LineRouteExt,"LineRouteExt");
addcell(query,this.LineToLineX,"LineToLineX");
addcell(query,this.LineToLineY,"LineToLineY");
addcell(query,this.LineToNodeX,"LineToNodeX");
addcell(query,this.LineToNodeY,"LineToNodeY");
addcell(query,this.PageBottomMargin,"PageBottomMargin");
addcell(query,this.PageHeight,"PageHeight");
addcell(query,this.PageLeftMargin,"PageLeftMargin");
addcell(query,this.PageLineJumpDirX,"PageLineJumpDirX");
addcell(query,this.PageLineJumpDirX,"PageLineJumpDirX");
addcell(query,this.PageLineJumpDirY,"PageLineJumpDirY");
addcell(query,this.PageLineJumpDirY,"PageLineJumpDirY");
addcell(query,this.PageRightMargin,"PageRightMargin");
addcell(query,this.PageScale,"PageScale");
addcell(query,this.PageShapeSplit,"PageShapeSplit");
addcell(query,this.PageShapeSplit,"PageShapeSplit");
addcell(query,this.PageTopMargin,"PageTopMargin");
addcell(query,this.PageWidth,"PageWidth");
addcell(query,this.PaperKind,"PaperKind");
addcell(query,this.PaperSource,"PaperSource");
addcell(query,this.PlaceDepth,"PlaceDepth");
addcell(query,this.PlaceFlip,"PlaceFlip");
addcell(query,this.PlaceStyle,"PlaceStyle");
addcell(query,this.PlowCode,"PlowCode");
addcell(query,this.PrintGrid,"PrintGrid");
addcell(query,this.PrintPageOrientation,"PrintPageOrientation");
addcell(query,this.ResizePage,"ResizePage");
addcell(query,this.RouteStyle,"RouteStyle");
addcell(query,this.ScaleX,"ScaleX");
addcell(query,this.ScaleY,"ScaleY");
addcell(query,this.ShdwObliqueAngle,"ShdwObliqueAngle");
addcell(query,this.ShdwOffsetX,"ShdwOffsetX");
addcell(query,this.ShdwOffsetY,"ShdwOffsetY");
addcell(query,this.ShdwScaleFactor,"ShdwScaleFactor");
addcell(query,this.ShdwType,"ShdwType");
addcell(query,this.UIVisibility,"UIVisibility");
addcell(query,this.XGridDensity,"XGridDensity");
addcell(query,this.XGridOrigin,"XGridOrigin");
addcell(query,this.XGridSpacing,"XGridSpacing");
addcell(query,this.XRulerDensity,"XRulerDensity");
addcell(query,this.XRulerOrigin,"XRulerOrigin");
addcell(query,this.YGridDensity,"YGridDensity");
addcell(query,this.YGridOrigin,"YGridOrigin");
addcell(query,this.YGridSpacing,"YGridSpacing");
addcell(query,this.YRulerDensity,"YRulerDensity");
addcell(query,this.YRulerOrigin,"YRulerOrigin");


            var page = scriptingsession.Page.Get();
            var target_shapeids = new[] { page.ID };

            this.WriteVerboseEx("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerboseEx("Start Query");

            var dt = VisioPSUtil.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, page);

            this.WriteObject(dt);
            this.WriteVerboseEx("End Query");
        }

        private void addcell(VisioAutomation.ShapeSheet.Query.CellQuery q, bool b, string name)
        {
            var dic = this.GetPageCellDictionary();
            if (b)
            {
                q.Columns.Add(dic[name], name);
            }
        }

        private static Dictionary<string, VA.ShapeSheet.SRC> dic_cellname_to_src;


        private Dictionary<string, VA.ShapeSheet.SRC> GetPageCellDictionary()
        {
            if (dic_cellname_to_src == null)
            {
                dic_cellname_to_src = new Dictionary<string, VA.ShapeSheet.SRC>();
                dic_cellname_to_src["PageBottomMargin"] = VA.ShapeSheet.SRCConstants.PageBottomMargin;
                dic_cellname_to_src["PageHeight"] = VA.ShapeSheet.SRCConstants.PageHeight;
                dic_cellname_to_src["PageLeftMargin"] = VA.ShapeSheet.SRCConstants.PageLeftMargin;
                dic_cellname_to_src["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;
                dic_cellname_to_src["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;
                dic_cellname_to_src["PageRightMargin"] = VA.ShapeSheet.SRCConstants.PageRightMargin;
                dic_cellname_to_src["PageScale"] = VA.ShapeSheet.SRCConstants.PageScale;
                dic_cellname_to_src["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;
                dic_cellname_to_src["PageTopMargin"] = VA.ShapeSheet.SRCConstants.PageTopMargin;
                dic_cellname_to_src["PageWidth"] = VA.ShapeSheet.SRCConstants.PageWidth;
                dic_cellname_to_src["CenterX"] = VA.ShapeSheet.SRCConstants.CenterX;
                dic_cellname_to_src["CenterY"] = VA.ShapeSheet.SRCConstants.CenterY;
                dic_cellname_to_src["PaperKind"] = VA.ShapeSheet.SRCConstants.PaperKind;
                dic_cellname_to_src["PrintGrid"] = VA.ShapeSheet.SRCConstants.PrintGrid;
                dic_cellname_to_src["PrintPageOrientation"] = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
                dic_cellname_to_src["ScaleX"] = VA.ShapeSheet.SRCConstants.ScaleX;
                dic_cellname_to_src["ScaleY"] = VA.ShapeSheet.SRCConstants.ScaleY;
                dic_cellname_to_src["PaperSource"] = VA.ShapeSheet.SRCConstants.PaperSource;
                dic_cellname_to_src["DrawingScale"] = VA.ShapeSheet.SRCConstants.DrawingScale;
                dic_cellname_to_src["DrawingScaleType"] = VA.ShapeSheet.SRCConstants.DrawingScaleType;
                dic_cellname_to_src["DrawingSizeType"] = VA.ShapeSheet.SRCConstants.DrawingSizeType;
                dic_cellname_to_src["InhibitSnap"] = VA.ShapeSheet.SRCConstants.InhibitSnap;
                dic_cellname_to_src["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                dic_cellname_to_src["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                dic_cellname_to_src["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                dic_cellname_to_src["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                dic_cellname_to_src["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                dic_cellname_to_src["UIVisibility"] = VA.ShapeSheet.SRCConstants.UIVisibility;
                dic_cellname_to_src["XGridDensity"] = VA.ShapeSheet.SRCConstants.XGridDensity;
                dic_cellname_to_src["XGridOrigin"] = VA.ShapeSheet.SRCConstants.XGridOrigin;
                dic_cellname_to_src["XGridSpacing"] = VA.ShapeSheet.SRCConstants.XGridSpacing;
                dic_cellname_to_src["XRulerDensity"] = VA.ShapeSheet.SRCConstants.XRulerDensity;
                dic_cellname_to_src["XRulerOrigin"] = VA.ShapeSheet.SRCConstants.XRulerOrigin;
                dic_cellname_to_src["YGridDensity"] = VA.ShapeSheet.SRCConstants.YGridDensity;
                dic_cellname_to_src["YGridOrigin"] = VA.ShapeSheet.SRCConstants.YGridOrigin;
                dic_cellname_to_src["YGridSpacing"] = VA.ShapeSheet.SRCConstants.YGridSpacing;
                dic_cellname_to_src["YRulerDensity"] = VA.ShapeSheet.SRCConstants.YRulerDensity;
                dic_cellname_to_src["YRulerOrigin"] = VA.ShapeSheet.SRCConstants.YRulerOrigin;
                dic_cellname_to_src["AvenueSizeX"] = VA.ShapeSheet.SRCConstants.AvenueSizeX;
                dic_cellname_to_src["AvenueSizeY"] = VA.ShapeSheet.SRCConstants.AvenueSizeY;
                dic_cellname_to_src["BlockSizeX"] = VA.ShapeSheet.SRCConstants.BlockSizeX;
                dic_cellname_to_src["BlockSizeY"] = VA.ShapeSheet.SRCConstants.BlockSizeY;
                dic_cellname_to_src["CtrlAsInput"] = VA.ShapeSheet.SRCConstants.CtrlAsInput;
                dic_cellname_to_src["DynamicsOff"] = VA.ShapeSheet.SRCConstants.DynamicsOff;
                dic_cellname_to_src["EnableGrid"] = VA.ShapeSheet.SRCConstants.EnableGrid;
                dic_cellname_to_src["LineAdjustFrom"] = VA.ShapeSheet.SRCConstants.LineAdjustFrom;
                dic_cellname_to_src["LineAdjustTo"] = VA.ShapeSheet.SRCConstants.LineAdjustTo;
                dic_cellname_to_src["LineJumpCode"] = VA.ShapeSheet.SRCConstants.LineJumpCode;
                dic_cellname_to_src["LineJumpFactorX"] = VA.ShapeSheet.SRCConstants.LineJumpFactorX;
                dic_cellname_to_src["LineJumpFactorY"] = VA.ShapeSheet.SRCConstants.LineJumpFactorY;
                dic_cellname_to_src["LineJumpStyle"] = VA.ShapeSheet.SRCConstants.LineJumpStyle;
                dic_cellname_to_src["LineRouteExt"] = VA.ShapeSheet.SRCConstants.LineRouteExt;
                dic_cellname_to_src["LineToLineX"] = VA.ShapeSheet.SRCConstants.LineToLineX;
                dic_cellname_to_src["LineToLineY"] = VA.ShapeSheet.SRCConstants.LineToLineY;
                dic_cellname_to_src["LineToNodeX"] = VA.ShapeSheet.SRCConstants.LineToNodeX;
                dic_cellname_to_src["LineToNodeY"] = VA.ShapeSheet.SRCConstants.LineToNodeY;
                dic_cellname_to_src["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;
                dic_cellname_to_src["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;
                dic_cellname_to_src["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;
                dic_cellname_to_src["PlaceDepth"] = VA.ShapeSheet.SRCConstants.PlaceDepth;
                dic_cellname_to_src["PlaceFlip"] = VA.ShapeSheet.SRCConstants.PlaceFlip;
                dic_cellname_to_src["PlaceStyle"] = VA.ShapeSheet.SRCConstants.PlaceStyle;
                dic_cellname_to_src["PlowCode"] = VA.ShapeSheet.SRCConstants.PlowCode;
                dic_cellname_to_src["ResizePage"] = VA.ShapeSheet.SRCConstants.ResizePage;
                dic_cellname_to_src["RouteStyle"] = VA.ShapeSheet.SRCConstants.RouteStyle;
                dic_cellname_to_src["AvoidPageBreaks"] = VA.ShapeSheet.SRCConstants.AvoidPageBreaks;
                dic_cellname_to_src["DrawingResizeType"] = VA.ShapeSheet.SRCConstants.DrawingResizeType;
            }
            return dic_cellname_to_src;
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