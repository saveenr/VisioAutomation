using System.Collections.Generic;
using System.Linq;
using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static class SRCConstants
    {
        // Actions
        public static SRC Actions_Action { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionAction); } }
        public static SRC Actions_BeginGroup { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionBeginGroup); } }
        public static SRC Actions_ButtonFace { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionButtonFace); } }
        public static SRC Actions_Checked { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionChecked); } }
        public static SRC Actions_Disabled { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionDisabled); } }
        public static SRC Actions_Invisible { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionInvisible); } }
        public static SRC Actions_Menu { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionMenu); } }
        public static SRC Actions_ReadOnly { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionReadOnly); } }
        public static SRC Actions_SortKey { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionSortKey); } }
        public static SRC Actions_TagName { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionTagName); } }
        public static SRC Actions_FlyoutChild { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionFlyoutChild); } } // new for visio 2010

        // Alignment
        public static SRC AlignBottom { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignBottom); } }
        public static SRC AlignCenter { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignCenter); } }
        public static SRC AlignLeft { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignLeft); } }
        public static SRC AlignMiddle { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignMiddle); } }
        public static SRC AlignRight { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignRight); } }
        public static SRC AlignTop { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignTop); } }

        // Annotation
        public static SRC Annotation_Comment { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationComment); } }
        public static SRC Annotation_Date { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationDate); } }
        public static SRC Annotation_LangID { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationLangID); } }
        public static SRC Annotation_MarkerIndex { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationMarkerIndex); } }
        public static SRC Annotation_X { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationX); } }
        public static SRC Annotation_Y { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationY); } }

        // Character
        public static SRC CharAsianFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterAsianFont); } }
        public static SRC CharCase { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterCase); } }
        public static SRC CharColor { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColor); } }
        public static SRC CharComplexScriptFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptFont); } }
        public static SRC CharComplexScriptSize { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptSize); } }
        public static SRC CharDoubleStrikethrough { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDoubleStrikethrough); } }
        public static SRC CharDblUnderline { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDblUnderline); } }
        public static SRC CharFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFont); } }
        public static SRC CharLangID { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLangID); } }
        public static SRC CharLocale { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocale); } }
        public static SRC CharLocalizeFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocalizeFont); } }
        public static SRC CharOverline { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterOverline); } }
        public static SRC CharPerpendicular { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPerpendicular); } }
        public static SRC CharPos { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPos); } }
        public static SRC CharRTLText { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterRTLText); } }
        public static SRC CharFontScale { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFontScale); } }
        public static SRC CharSize { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterSize); } }
        public static SRC CharLetterspace { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLetterspace); } }
        public static SRC CharStrikethru { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStrikethru); } }
        public static SRC CharStyle { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStyle); } }
        public static SRC CharColorTrans { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColorTrans); } }
        public static SRC CharUseVertical { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterUseVertical); } }

        // Connections
        public static SRC Connections_D { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctD); } }
        public static SRC Connections_DirX { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirX); } }
        public static SRC Connections_DirY { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirY); } }
        public static SRC Connections_Type { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctType); } }
        public static SRC Connections_X { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visX); } }
        public static SRC Connections_Y { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visY); } }

        // Controls
        public static SRC Controls_CanGlue { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlGlue); } }
        public static SRC Controls_Tip { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlTip); } }
        public static SRC Controls_XCon { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXCon); } }
        public static SRC Controls_X { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlX); } }
        public static SRC Controls_XDyn { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXDyn); } }
        public static SRC Controls_YCon { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYCon); } }
        public static SRC Controls_Y { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlY); } }
        public static SRC Controls_YDyn { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYDyn); } }

        // Document Properties
        public static SRC AddMarkup { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocAddMarkup); } }
        public static SRC DocLangID { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLangID); } }
        public static SRC LockPreview { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLockPreview); } }
        public static SRC OutputFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocOutputFormat); } }
        public static SRC PreviewQuality { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewQuality); } }
        public static SRC PreviewScope { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewScope); } }
        public static SRC ViewMarkup { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocViewMarkup); } }
        
        // Events
        public static SRC EventDblClick { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDblClick); } }
        public static SRC EventDrop { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDrop); } }
        public static SRC EventMultiDrop { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellMultiDrop); } }
        public static SRC EventXFMod { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellXFMod); } }
        public static SRC TheText { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellTheText); } }

        // ForeignImageInfo
        public static SRC ImgHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgHeight); } }
        public static SRC ImgOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetX); } }
        public static SRC ImgOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetY); } }
        public static SRC ImgWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgWidth); } }

        // Geometry
        public static SRC Geometry_A { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visBow); } }
        public static SRC Geometry_B { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visControlX); } }
        public static SRC Geometry_C { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visEccentricityAngle); } }
        public static SRC Geometry_D { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visAspectRatio); } }
        public static SRC Geometry_E { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visNURBSData); } }
        public static SRC Geometry_X { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visX); } }
        public static SRC Geometry_Y { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visY); } }
        public static SRC Geometry_NoFill { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoFill); } }
        public static SRC Geometry_NoLine { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoLine); } }
        public static SRC Geometry_NoShow { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoShow); } }
        public static SRC Geometry_NoSnap { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoSnap); } }
        public static SRC Geometry_NoQuickDrag { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoQuickDrag); } }

        // Fill Format
        public static SRC FillBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgnd); } }
        public static SRC FillBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgndTrans); } }
        public static SRC FillForegnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd); } }
        public static SRC FillForegndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegndTrans); } }
        public static SRC FillPattern { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillPattern); } }
        public static SRC ShapeShdwObliqueAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwObliqueAngle); } }
        public static SRC ShapeShdwOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetX); } }
        public static SRC ShapeShdwOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetY); } }
        public static SRC ShapeShdwScaleFactor { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwScaleFactor); } }
        public static SRC ShapeShdwType { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwType); } }
        public static SRC ShdwBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgnd); } }
        public static SRC ShdwBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgndTrans); } }
        public static SRC ShdwForegnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegnd); } }
        public static SRC ShdwForegndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegndTrans); } }
        public static SRC ShdwPattern { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwPattern); } }

        // GlueInfo
        public static SRC BegTrigger { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visBegTrigger); } }
        public static SRC EndTrigger { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visEndTrigger); } }
        public static SRC GlueType { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visGlueType); } }
        public static SRC WalkPreference { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visWalkPref); } }

        // GroupProperties
        public static SRC DisplayMode { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDisplayMode); } }
        public static SRC DontMoveChildren { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDontMoveChildren); } }
        public static SRC IsDropTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsDropTarget); } }
        public static SRC IsSnapTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsSnapTarget); } }
        public static SRC IsTextEditTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsTextEditTarget); } }
        public static SRC SelectMode { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupSelectMode); } }

        // Hyperlinks
        public static SRC Hyperlink_Address { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkAddress); } }
        public static SRC Hyperlink_Default { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDefault); } }
        public static SRC Hyperlink_Description { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDescription); } }
        public static SRC Hyperlink_ExtraInfo { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkExtraInfo); } }
        public static SRC Hyperlink_Frame { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkFrame); } }
        public static SRC Hyperlink_Invisible { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkInvisible); } }
        public static SRC Hyperlink_NewWindow { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkNewWin); } }
        public static SRC Hyperlink_SortKey { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSortKey); } }
        public static SRC Hyperlink_SubAddress { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSubAddress); } }

        // Image Properties
        public static SRC Blur { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBlur); } }
        public static SRC Brightness { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBrightness); } }
        public static SRC Contrast { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageContrast); } }
        public static SRC Denoise { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageDenoise); } }
        public static SRC Gamma { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageGamma); } }
        public static SRC Sharpen { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageSharpen); } }
        public static SRC Transparency { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageTransparency); } }

        // Line format
        public static SRC BeginArrow { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrow); } }
        public static SRC BeginArrowSize { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrowSize); } }
        public static SRC EndArrow { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrow); } }
        public static SRC EndArrowSize { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrowSize); } }
        public static SRC LineCap { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndCap); } }
        public static SRC LineColor { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColor); } }
        public static SRC LineColorTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColorTrans); } }
        public static SRC LinePattern { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLinePattern); } }
        public static SRC LineWeight { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineWeight); } }
        public static SRC Rounding { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineRounding); } }

        // Miscellaneous
        public static SRC Calendar { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjCalendar); } }
        public static SRC Comment { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visComment); } }
        public static SRC DropOnPageScale { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjDropOnPageScale); } }
        public static SRC DynFeedback { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDynFeedback); } }
        public static SRC IsDropSource { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDropSource); } }
        public static SRC LangID { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLangID); } }
        public static SRC LocalizeMerge { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLocalizeMerge); } }
        public static SRC NoAlignBox { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoAlignBox); } }
        public static SRC NoCtlHandles { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoCtlHandles); } }
        public static SRC NoLiveDynamics { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoLiveDynamics); } }
        public static SRC NonPrinting { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNonPrinting); } }
        public static SRC NoObjHandles { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoObjHandles); } }
        public static SRC ObjType { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visLOFlags); } }
        public static SRC UpdateAlignBox { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visUpdateAlignBox); } }
        public static SRC HideText { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visHideText); } }

        // 1d endpoints
        public static SRC BeginX { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginX); } }
        public static SRC BeginY { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginY); } }
        public static SRC EndX { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndX); } }
        public static SRC EndY { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndY); } }

        // page layout
        public static SRC AvenueSizeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeX); } }
        public static SRC AvenueSizeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeY); } }
        public static SRC BlockSizeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeX); } }
        public static SRC BlockSizeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeY); } }
        public static SRC CtrlAsInput { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOCtrlAsInput); } }
        public static SRC DynamicsOff { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLODynamicsOff); } }
        public static SRC EnableGrid { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOEnableGrid); } }
        public static SRC LineAdjustFrom { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustFrom); } }
        public static SRC LineAdjustTo { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustTo); } }
        public static SRC LineJumpCode { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpCode); } }
        public static SRC LineJumpFactorX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorX); } }
        public static SRC LineJumpFactorY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorY); } }
        public static SRC LineJumpStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpStyle); } }
        public static SRC LineRouteExt { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineRouteExt); } }
        public static SRC LineToLineX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineX); } }
        public static SRC LineToLineY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineY); } }
        public static SRC LineToNodeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeX); } }
        public static SRC LineToNodeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeY); } }
        public static SRC PageLineJumpDirX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirX); } }
        public static SRC PageLineJumpDirY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirY); } }
        public static SRC PageShapeSplit { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOSplit); } }
        public static SRC PlaceDepth { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceDepth); } }
        public static SRC PlaceFlip { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceFlip); } }
        public static SRC PlaceStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceStyle); } }
        public static SRC PlowCode { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlowCode); } }
        public static SRC ResizePage { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOResizePage); } }
        public static SRC RouteStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLORouteStyle); } }

        public static SRC AvoidPageBreaks { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvoidPageBreaks); } } // new in Visio 2010

        // print properties
        public static SRC PageLeftMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesLeftMargin); } }
        public static SRC CenterX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterX); } }
        public static SRC CenterY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterY); } }
        public static SRC OnPage { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesOnPage); } }
        public static SRC PageBottomMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesBottomMargin); } }
        public static SRC PageRightMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesRightMargin); } }
        public static SRC PagesX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesX); } }
        public static SRC PagesY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesY); } }
        public static SRC PageTopMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesTopMargin); } }
        public static SRC PaperKind { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperKind); } }
        public static SRC PrintGrid { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPrintGrid); } }
        public static SRC PrintPageOrientation { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPageOrientation); } }
        public static SRC ScaleX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleX); } }
        public static SRC ScaleY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleY); } }
        public static SRC PaperSource { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperSource); } }

        // page properties
        public static SRC DrawingScale { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawingScale); } }
        public static SRC DrawingScaleType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawScaleType); } }
        public static SRC DrawingSizeType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawSizeType); } }
        public static SRC InhibitSnap { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageInhibitSnap); } }
        public static SRC PageHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageHeight); } }
        public static SRC PageScale { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageScale); } }
        public static SRC PageWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageWidth); } }
        public static SRC ShdwObliqueAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwObliqueAngle); } }
        public static SRC ShdwOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetX); } }
        public static SRC ShdwOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetY); } }
        public static SRC ShdwScaleFactor { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwScaleFactor); } }
        public static SRC ShdwType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwType); } }
        public static SRC UIVisibility { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageUIVisibility); } }
        public static SRC DrawingResizeType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawResizeType); } } // new in Visio 2010

        // paragraph
        public static SRC Para_Bullet { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletIndex); } }
        public static SRC Para_BulletFont { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFont); } }
        public static SRC Para_BulletFontSize { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFontSize); } }
        public static SRC Para_BulletStr { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletString); } }
        public static SRC Para_Flags { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visFlags); } }
        public static SRC Para_HorzAlign { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visHorzAlign); } }
        public static SRC Para_IndFirst { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentFirst); } }
        public static SRC Para_IndLeft { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentLeft); } }
        public static SRC Para_IndRight { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentRight); } }
        public static SRC Para_LocalizeBulletFont { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visLocalizeBulletFont); } }
        public static SRC Para_SpAfter { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceAfter); } }
        public static SRC Para_SpBefore { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceBefore); } }
        public static SRC Para_SpLine { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceLine); } }
        public static SRC Para_TextPosAfterBullet { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visTextPosAfterBullet); } }

        // protection
        public static SRC LockAspect { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockAspect); } }
        public static SRC LockBegin { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockBegin); } }
        public static SRC LockCalcWH { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCalcWH); } }
        public static SRC LockCrop { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCrop); } }
        public static SRC LockCustProp { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCustProp); } }
        public static SRC LockDelete { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockDelete); } }
        public static SRC LockEnd { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockEnd); } }
        public static SRC LockFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFormat); } }
        public static SRC LockFromGroupFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFromGroupFormat); } }
        public static SRC LockGroup { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockGroup); } }
        public static SRC LockHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockHeight); } }
        public static SRC LockMoveX { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveX); } }
        public static SRC LockMoveY { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveY); } }
        public static SRC LockRotate { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockRotate); } }
        public static SRC LockSelect { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockSelect); } }
        public static SRC LockTextEdit { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockTextEdit); } }
        public static SRC LockThemeColors { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeColors); } }
        public static SRC LockThemeEffects { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeEffects); } }
        public static SRC LockVtxEdit { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockVtxEdit); } }
        public static SRC LockWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockWidth); } }

        // ruler and grid
        public static SRC XGridDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridDensity); } }
        public static SRC XGridOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridOrigin); } }
        public static SRC XGridSpacing { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridSpacing); } }
        public static SRC XRulerDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerDensity); } }
        public static SRC XRulerOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerOrigin); } }
        public static SRC YGridDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridDensity); } }
        public static SRC YGridOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridOrigin); } }
        public static SRC YGridSpacing { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridSpacing); } }
        public static SRC YRulerDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerDensity); } }
        public static SRC YRulerOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerOrigin); } }

        // Shape Tranform
        public static SRC Angle { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormAngle); } }
        public static SRC FlipX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipX); } }
        public static SRC FlipY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipY); } }
        public static SRC Height { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormHeight); } }
        public static SRC LocPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinX); } }
        public static SRC LocPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinY); } }
        public static SRC PinX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinX); } }
        public static SRC PinY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinY); } }
        public static SRC ResizeMode { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormResizeMode); } }
        public static SRC Width { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormWidth); } }

        // reviewer
        public static SRC Reviewer_Color { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerColor); } }
        public static SRC Reviewer_Initials { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerInitials); } }
        public static SRC Reviewer_Name { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerName); } }

        // shape data
        public static SRC Prop_SortKey { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsSortKey); } }
        public static SRC Prop_Ask { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsAsk); } }
        public static SRC Prop_Calendar { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsCalendar); } }
        public static SRC Prop_Format { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsFormat); } }
        public static SRC Prop_Invisible { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsInvis); } }
        public static SRC Prop_Label { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLabel); } }
        public static SRC Prop_LangID { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLangID); } }
        public static SRC Prop_Prompt { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsPrompt); } }
        public static SRC Prop_Type { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsType); } }
        public static SRC Prop_Value { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsValue); } }

        // Layers
        public static SRC Layers_Active { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerActive); } }
        public static SRC Layers_Color { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColor); } }
        public static SRC Layers_Glue { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerGlue); } }
        public static SRC Layers_Locked { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerLock); } }
        public static SRC Layers_Print { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visDocPreviewScope); } }
        public static SRC Layers_Snap { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerSnap); } }
        public static SRC Layers_ColorTrans { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColorTrans); } }
        public static SRC Layers_Visible { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerVisible); } }

        //text transform
        public static SRC TxtAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormAngle); } }
        public static SRC TxtHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormHeight); } }
        public static SRC TxtLocPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinX); } }
        public static SRC TxtLocPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinY); } }
        public static SRC TxtPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinX); } }
        public static SRC TxtPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinY); } }
        public static SRC TxtWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormWidth); } }

        // user defined cells
        public static SRC User_Prompt { get { return new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserPrompt); } }
        public static SRC User_Value { get { return new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserValue); } }

        // Fields
        public static SRC Fields_Calendar { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCalendar); } }
        public static SRC Fields_Format { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldFormat); } }
        public static SRC Fields_ObjectKind { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldObjectKind); } }
        public static SRC Fields_Type { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldType); } }
        public static SRC Fields_UICat { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICategory); } }
        public static SRC Fields_UICod { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICode); } }
        public static SRC Fields_UIFmt { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUIFormat); } }
        public static SRC Fields_Value { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCell); } }

        // text block format
        public static SRC BottomMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBottomMargin); } }
        public static SRC DefaultTabStop { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDefaultTabStop); } }
        public static SRC LeftMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkLeftMargin); } }
        public static SRC RightMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkRightMargin); } }
        public static SRC TextBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgnd); } }
        public static SRC TextBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgndTrans); } }
        public static SRC TextDirection { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDirection); } }
        public static SRC TopMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkTopMargin); } }
        public static SRC VerticalAlign { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkVerticalAlign); } }

        // Action tags
        public static SRC SmartTags_ButtonFace { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagButtonFace); } }
        public static SRC SmartTags_Description { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDescription); } }
        public static SRC SmartTags_Disabled { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisabled); } }
        public static SRC SmartTags_DisplayMode { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisplayMode); } }
        public static SRC SmartTags_TagName { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagName); } }
        public static SRC SmartTags_X { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagX); } }
        public static SRC SmartTags_XJustify { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagXJustify); } }
        public static SRC SmartTags_Y { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagY); } }
        public static SRC SmartTags_YJustify { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagYJustify); } }

        // style
        public static SRC EnableFillProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesFill); } }
        public static SRC EnableLineProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesLine); } }
        public static SRC EnableTextProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesText); } }

        //tabs
        public static SRC Tabs_Alignment { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabAlign); } }
        public static SRC Tabs_Position { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabPos); } }
        public static SRC Tabs_StopCount { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabStopCount); } }

        // shape layout
        public static SRC ConFixedCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOConFixedCode); } }
        public static SRC ConLineJumpCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpCode); } }
        public static SRC ConLineJumpDirX { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirX); } }
        public static SRC ConLineJumpDirY { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirY); } }
        public static SRC ConLineJumpStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpStyle); } }
        public static SRC ConLineRouteExt { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOLineRouteExt); } }
        public static SRC ShapeFixedCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOFixedCode); } }
        public static SRC ShapePermeablePlace { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermeablePlace); } }
        public static SRC ShapePermeableX { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermX); } }
        public static SRC ShapePermeableY { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermY); } }
        public static SRC ShapePlaceFlip { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceFlip); } }
        public static SRC ShapePlaceStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceStyle); } }
        public static SRC ShapePlowCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlowCode); } }
        public static SRC ShapeRouteStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORouteStyle); } }
        public static SRC ShapeSplit { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplit); } }
        public static SRC ShapeSplittable { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplittable); } }
        public static SRC DisplayLevel { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLODisplayLevel); } } // new in Visio 2010
        public static SRC Relationships { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORelationships); } } // new in Visio 2010

        public static Dictionary<string, SRC> GetSRCDictionary()
        {
            var srcconstants_t = typeof(SRCConstants);

            var binding_flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static;

            // find all the static properties that return SRC types
            var src_type = typeof (SRC);
            var props = srcconstants_t.GetProperties(binding_flags)
                .Where(p => p.PropertyType == src_type);

            var fields_name_to_value = new Dictionary<string, SRC>();
            foreach (var propinfo in props)
            {
                var src = (SRC) propinfo.GetValue(null,null);
                var name = propinfo.Name;
                fields_name_to_value[name] = src;
            }

            return fields_name_to_value;
        }
    }
}