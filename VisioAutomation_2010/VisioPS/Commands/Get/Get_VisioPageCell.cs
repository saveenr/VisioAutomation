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
        [SMA.Parameter(Mandatory = true,Position=0)]
        [SMA.ValidateSet(
            "PageBottomMargin", "PageHeight", "PageLeftMargin", "PageLineJumpDirX", "PageLineJumpDirY", 
            "PageRightMargin", "PageScale", "PageShapeSplit", "PageTopMargin", "PageWidth", "CenterX", "CenterY", 
            "PaperKind", "PrintGrid", "PrintPageOrientation", "ScaleX", "ScaleY", "PaperSource", 
            "DrawingScale", "DrawingScaleType", "DrawingSizeType", "InhibitSnap", "ShdwObliqueAngle", 
            "ShdwOffsetX", "ShdwOffsetY", "ShdwScaleFactor", "ShdwType", "UIVisibility", "XGridDensity", 
            "XGridOrigin", "XGridSpacing", "XRulerDensity", "XRulerOrigin", "YGridDensity", "YGridOrigin", "YGridSpacing", 
            "YRulerDensity", "YRulerOrigin", "AvenueSizeX", "AvenueSizeY", "BlockSizeX", "BlockSizeY", "CtrlAsInput", 
            "DynamicsOff", "EnableGrid", "LineAdjustFrom", "LineAdjustTo", "LineJumpCode", "LineJumpFactorX", "LineJumpFactorY", 
            "LineJumpStyle", "LineRouteExt", "LineToLineX", "LineToLineY", "LineToNodeX", "LineToNodeY", "PageLineJumpDirX", 
            "PageLineJumpDirY", "PageShapeSplit", "PlaceDepth", "PlaceFlip", "PlaceStyle", "PlowCode", "ResizePage", 
            "RouteStyle", "AvoidPageBreaks", "DrawingResizeType")]

        public string[] Cells { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var query = new VA.ShapeSheet.Query.CellQuery();

            var dic = GetPageCellDictionary();
            foreach (var cell in this.Cells)
            {
                query.AddColumn(dic[cell], cell);   
            }

            var page = scriptingsession.Page.Get();
            var target_shapeids = new[] { page.ID };

            this.WriteVerboseEx("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerboseEx("Start Query");

            var dt = VisioPSUtil.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, page);

            this.WriteObject(dt);
            this.WriteVerboseEx("End Query");
        }

        private static Dictionary<string, VA.ShapeSheet.SRC> dic_cellname_to_src;


        private Dictionary<string, VA.ShapeSheet.SRC> GetPageCellDictionary()
        {
            if (dic_cellname_to_src == null)
            {
                dic_cellname_to_src = new Dictionary<string, VA.ShapeSheet.SRC>(this.Cells.Count());
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
    }
}