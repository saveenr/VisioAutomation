using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public string[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)] 
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {

            if (this.Cells == null)
            {
                throw new System.ArgumentException("Cells");
            }

            if (this.Cells.Length < 1)
            {
                string msg = "Must provide at least one cell name";
                throw new System.ArgumentException(msg);
            }

            var map = GetShapeCellDictionary();
            var invalid_names = this.Cells.Where(cellname => !map.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                var names = map.GetNames();
                string valid_names = string.Join(",", names);
                this.WriteVerbose( "Valid Names: " + valid_names);
                string msg = "Invalid cell names: " + string.Join(",",invalid_names);
                throw new System.ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();

            var dic = GetShapeCellDictionary();
            Get_VisioPageCell.SetFromCellNames(query, this.Cells, dic);

            var surface = this.client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Cells: {0}", query.CellColumns.Count);

            this.WriteVerbose("Start Query");

            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);

            this.WriteVerbose("End Query");
        }

        private static CellMap map_name_to_cell;

        public static CellMap GetShapeCellDictionary()
        {
            if (map_name_to_cell == null)
            {
                map_name_to_cell = new CellMap();
                map_name_to_cell[VA.ShapeSheet.SRCConstants.Angle.Name] = VA.ShapeSheet.SRCConstants.Angle;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.BeginX.Name] = VA.ShapeSheet.SRCConstants.BeginX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.BeginY.Name] = VA.ShapeSheet.SRCConstants.BeginY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharCase.Name] = VA.ShapeSheet.SRCConstants.CharCase;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharColor.Name] = VA.ShapeSheet.SRCConstants.CharColor;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharColorTrans.Name] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharFont.Name] = VA.ShapeSheet.SRCConstants.CharFont;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharFontScale.Name] = VA.ShapeSheet.SRCConstants.CharFontScale;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharLetterspace.Name] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharSize.Name] = VA.ShapeSheet.SRCConstants.CharSize;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.CharStyle.Name] = VA.ShapeSheet.SRCConstants.CharStyle;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.EndX.Name] = VA.ShapeSheet.SRCConstants.EndX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.EndY.Name] = VA.ShapeSheet.SRCConstants.EndY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.FillBkgnd.Name] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.FillBkgndTrans.Name] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.FillForegnd.Name] = VA.ShapeSheet.SRCConstants.FillForegnd;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.FillForegndTrans.Name] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.FillPattern.Name] = VA.ShapeSheet.SRCConstants.FillPattern;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.Height.Name] = VA.ShapeSheet.SRCConstants.Height;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LineCap.Name] = VA.ShapeSheet.SRCConstants.LineCap;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LineColor.Name] = VA.ShapeSheet.SRCConstants.LineColor;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LinePattern.Name] = VA.ShapeSheet.SRCConstants.LinePattern;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LineWeight.Name] = VA.ShapeSheet.SRCConstants.LineWeight;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockAspect.Name] = VA.ShapeSheet.SRCConstants.LockAspect;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockBegin.Name] = VA.ShapeSheet.SRCConstants.LockBegin;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockCalcWH.Name] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockCrop.Name] = VA.ShapeSheet.SRCConstants.LockCrop;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockCustProp.Name] = VA.ShapeSheet.SRCConstants.LockCustProp;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockDelete.Name] = VA.ShapeSheet.SRCConstants.LockDelete;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockEnd.Name] = VA.ShapeSheet.SRCConstants.LockEnd;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockFormat.Name] = VA.ShapeSheet.SRCConstants.LockFormat;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockFromGroupFormat.Name] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockGroup.Name] = VA.ShapeSheet.SRCConstants.LockGroup;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockHeight.Name] = VA.ShapeSheet.SRCConstants.LockHeight;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockMoveX.Name] = VA.ShapeSheet.SRCConstants.LockMoveX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockMoveY.Name] = VA.ShapeSheet.SRCConstants.LockMoveY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockRotate.Name] = VA.ShapeSheet.SRCConstants.LockRotate;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockSelect.Name] = VA.ShapeSheet.SRCConstants.LockSelect;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockTextEdit.Name] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockThemeColors.Name] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockThemeEffects.Name] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockVtxEdit.Name] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LockWidth.Name] = VA.ShapeSheet.SRCConstants.LockWidth;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LocPinX.Name] = VA.ShapeSheet.SRCConstants.LocPinX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.LocPinY.Name] = VA.ShapeSheet.SRCConstants.LocPinY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.PinX.Name] = VA.ShapeSheet.SRCConstants.PinX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.PinY.Name] = VA.ShapeSheet.SRCConstants.PinY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.Rounding.Name] = VA.ShapeSheet.SRCConstants.Rounding;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.SelectMode.Name] = VA.ShapeSheet.SRCConstants.SelectMode;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwBkgnd.Name] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwBkgndTrans.Name] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwForegnd.Name] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwForegndTrans.Name] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwObliqueAngle.Name] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwOffsetX.Name] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwOffsetY.Name] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwPattern.Name] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwScaleFactor.Name] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.ShdwType.Name] = VA.ShapeSheet.SRCConstants.ShdwType;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtAngle.Name] = VA.ShapeSheet.SRCConstants.TxtAngle;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtHeight.Name] = VA.ShapeSheet.SRCConstants.TxtHeight;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtLocPinX.Name] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtLocPinY.Name] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtPinX.Name] = VA.ShapeSheet.SRCConstants.TxtPinX;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtPinY.Name] = VA.ShapeSheet.SRCConstants.TxtPinY;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.TxtWidth.Name] = VA.ShapeSheet.SRCConstants.TxtWidth;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.Width.Name] = VA.ShapeSheet.SRCConstants.Width;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.BeginArrow.Name] = VA.ShapeSheet.SRCConstants.BeginArrow;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.BeginArrowSize.Name] = VA.ShapeSheet.SRCConstants.BeginArrowSize;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.EndArrow.Name] = VA.ShapeSheet.SRCConstants.EndArrow;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.EndArrowSize.Name] = VA.ShapeSheet.SRCConstants.EndArrowSize;
                map_name_to_cell[VA.ShapeSheet.SRCConstants.HideText.Name] = VA.ShapeSheet.SRCConstants.HideText;
            }
            return map_name_to_cell;
        }
    }

    /*

Angle  
BeginArrow  
BeginArrowSize  
BeginX  
BeginY  
CharCase  
CharColor  
CharColorTransparency  
CharFont  
CharFontScale  
CharLetterspace  
CharSize  
CharStyle  
EndArrow  
EndArrowSize  
EndX  
EndY  
FillBkgnd  
FillBkgndTrans  
FillForegnd  
FillForegndTrans  
FillPattern  
Height  
HideText  
LineCap  
LineColor  
LinePattern  
LineWeight  
LockAspect  
LockBegin  
LockCalcWH  
LockCrop  
LockCustProp  
LockDelete  
LockEnd  
LockFormat  
LockFromGroupFormat  
LockGroup  
LockHeight  
LockMoveX  
LockMoveY  
LockRotate  
LockSelect  
LockTextEdit  
LockThemeColors  
LockThemeEffects  
LockVtxEdit  
LockWidth  
LocPinX  
LocPinY  
PinX  
PinY  
Rounding  
SelectMode  
ShdwBkgnd  
ShdwBkgndTrans  
ShdwForegnd  
ShdwForegndTrans  
ShdwObliqueAngle  
ShdwOffsetX  
ShdwOffsetY  
ShdwPattern  
ShdwScalefactor  
ShdwType  
TxtAngle  
TxtHeight  
TxtLocPinX  
TxtLocPinY  
TxtPinX  
TxtPinY  
TxtWidth  
Width  

     
     */
}