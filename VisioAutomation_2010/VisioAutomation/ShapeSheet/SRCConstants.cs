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
        public static SRC Actions_Action { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionAction, "Action"); } }
        public static SRC Actions_BeginGroup { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionBeginGroup, "BeginGroup"); } }
        public static SRC Actions_ButtonFace { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionButtonFace, "ButtonFace"); } }
        public static SRC Actions_Checked { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionChecked, "Checked"); } }
        public static SRC Actions_Disabled { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionDisabled, "Disabled"); } }
        public static SRC Actions_Invisible { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionInvisible, "Invisible"); } }
        public static SRC Actions_Menu { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionMenu, "Menu"); } }
        public static SRC Actions_ReadOnly { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionReadOnly, "ReadOnly"); } }
        public static SRC Actions_SortKey { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionSortKey, "SortKey"); } }
        public static SRC Actions_TagName { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionTagName, "TagName"); } }
        public static SRC Actions_FlyoutChild { get { return new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionFlyoutChild, "FlyoutChild"); } } // new for visio 2010

        // Alignment
        public static SRC AlignBottom { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignBottom, "AlignBottom"); } }
        public static SRC AlignCenter { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignCenter, "AlignCenter"); } }
        public static SRC AlignLeft { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignLeft, "AlignLeft"); } }
        public static SRC AlignMiddle { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignMiddle, "AlignMiddle"); } }
        public static SRC AlignRight { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignRight, "AlignRight"); } }
        public static SRC AlignTop { get { return new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignTop, "AlignTop"); } }

        // Annotation
        public static SRC Annotation_Comment { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationComment, "Comment"); } }
        public static SRC Annotation_Date { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationDate, "Date"); } }
        public static SRC Annotation_LangID { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationLangID, "LangID"); } }
        public static SRC Annotation_MarkerIndex { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationMarkerIndex, "MarkerIndex"); } }
        public static SRC Annotation_X { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationX, "X"); } }
        public static SRC Annotation_Y { get { return new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationY, "Y"); } }

        // Character
        public static SRC CharAsianFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterAsianFont, "CharAsianFont"); } }
        public static SRC CharCase { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterCase, "CharCase"); } }
        public static SRC CharColor { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColor, "CharColor"); } }
        public static SRC CharComplexScriptFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptFont, "CharComplexScriptFont"); } }
        public static SRC CharComplexScriptSize { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptSize, "CharComplexScriptSize"); } }
        public static SRC CharDoubleStrikethrough { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDoubleStrikethrough, "CharDoubleStrikethrough"); } }
        public static SRC CharDblUnderline { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDblUnderline, "CharDblUnderline"); } }
        public static SRC CharFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFont, "CharFont"); } }
        public static SRC CharLangID { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLangID, "CharLangID"); } }
        public static SRC CharLocale { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocale, "CharLocale"); } }
        public static SRC CharLocalizeFont { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocalizeFont, "CharLocalizeFont"); } }
        public static SRC CharOverline { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterOverline, "CharOverline"); } }
        public static SRC CharPerpendicular { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPerpendicular, "CharPerpendicular"); } }
        public static SRC CharPos { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPos, "CharPos"); } }
        public static SRC CharRTLText { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterRTLText, "CharRTLText"); } }
        public static SRC CharFontScale { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFontScale, "CharFontScale"); } }
        public static SRC CharSize { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterSize, "CharSize"); } }
        public static SRC CharLetterspace { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLetterspace, "CharLetterspace"); } }
        public static SRC CharStrikethru { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStrikethru, "CharStrikethru"); } }
        public static SRC CharStyle { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStyle, "CharStyle"); } }
        public static SRC CharColorTrans { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColorTrans, "CharColorTrans"); } }
        public static SRC CharUseVertical { get { return new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterUseVertical, "CharUseVertical"); } }

        // Connections
        public static SRC Connections_D { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctD, "D"); } }
        public static SRC Connections_DirX { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirX, "DirX"); } }
        public static SRC Connections_DirY { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirY, "DirY"); } }
        public static SRC Connections_Type { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctType, "Type"); } }
        public static SRC Connections_X { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visX, "X"); } }
        public static SRC Connections_Y { get { return new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visY, "Y"); } }

        // Controls
        public static SRC Controls_CanGlue { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlGlue, "CanGlue"); } }
        public static SRC Controls_Tip { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlTip, "Tip"); } }
        public static SRC Controls_XCon { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXCon, "XCon"); } }
        public static SRC Controls_X { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlX, "X"); } }
        public static SRC Controls_XDyn { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXDyn, "XDyn"); } }
        public static SRC Controls_YCon { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYCon, "YCon"); } }
        public static SRC Controls_Y { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlY, "Y"); } }
        public static SRC Controls_YDyn { get { return new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYDyn, "YDyn"); } }

        // Document Properties
        public static SRC AddMarkup { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocAddMarkup, "AddMarkup"); } }
        public static SRC DocLangID { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLangID, "DocLangID"); } }
        public static SRC LockPreview { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLockPreview, "LockPreview"); } }
        public static SRC OutputFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocOutputFormat, "OutputFormat"); } }
        public static SRC PreviewQuality { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewQuality, "PreviewQuality"); } }
        public static SRC PreviewScope { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewScope, "PreviewScope"); } }
        public static SRC ViewMarkup { get { return new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocViewMarkup, "ViewMarkup"); } }

        // Events
        public static SRC EventDblClick { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDblClick, "EventDblClick"); } }
        public static SRC EventDrop { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDrop, "EventDrop"); } }
        public static SRC EventMultiDrop { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellMultiDrop, "EventMultiDrop"); } }
        public static SRC EventXFMod { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellXFMod, "EventXFMod"); } }
        public static SRC TheText { get { return new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellTheText, "TheText"); } }

        // ForeignImageInfo
        public static SRC ImgHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgHeight, "ImgHeight"); } }
        public static SRC ImgOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetX, "ImgOffsetX"); } }
        public static SRC ImgOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetY, "ImgOffsetY"); } }
        public static SRC ImgWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgWidth, "ImgWidth"); } }

        // Geometry
        public static SRC Geometry_A { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visBow, "A"); } }
        public static SRC Geometry_B { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visControlX, "B"); } }
        public static SRC Geometry_C { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visEccentricityAngle, "C"); } }
        public static SRC Geometry_D { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visAspectRatio, "D"); } }
        public static SRC Geometry_E { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visNURBSData, "E"); } }
        public static SRC Geometry_X { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visX, "X"); } }
        public static SRC Geometry_Y { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visY, "Y"); } }
        public static SRC Geometry_NoFill { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoFill, "NoFill"); } }
        public static SRC Geometry_NoLine { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoLine, "NoLine"); } }
        public static SRC Geometry_NoShow { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoShow, "NoShow"); } }
        public static SRC Geometry_NoSnap { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoSnap, "NoSnap"); } }
        public static SRC Geometry_NoQuickDrag { get { return new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoQuickDrag, "NoQuickDrag"); } }

        // Fill Format
        public static SRC FillBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgnd, "FillBkgnd"); } }
        public static SRC FillBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgndTrans, "FillBkgndTrans"); } }
        public static SRC FillForegnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd, "FillForegnd"); } }
        public static SRC FillForegndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegndTrans, "FillForegndTrans"); } }
        public static SRC FillPattern { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillPattern, "FillPattern"); } }
        public static SRC ShapeShdwObliqueAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwObliqueAngle, "ShapeShdwObliqueAngle"); } }
        public static SRC ShapeShdwOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetX, "ShapeShdwOffsetX"); } }
        public static SRC ShapeShdwOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetY, "ShapeShdwOffsetY"); } }
        public static SRC ShapeShdwScaleFactor { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwScaleFactor, "ShapeShdwScaleFactor"); } }
        public static SRC ShapeShdwType { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwType, "ShapeShdwType"); } }
        public static SRC ShdwBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgnd, "ShdwBkgnd"); } }
        public static SRC ShdwBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgndTrans, "ShdwBkgndTrans"); } }
        public static SRC ShdwForegnd { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegnd, "ShdwForegnd"); } }
        public static SRC ShdwForegndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegndTrans, "ShdwForegndTrans"); } }
        public static SRC ShdwPattern { get { return new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwPattern, "ShdwPattern"); } }

        // GlueInfo
        public static SRC BegTrigger { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visBegTrigger, "BegTrigger"); } }
        public static SRC EndTrigger { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visEndTrigger, "EndTrigger"); } }
        public static SRC GlueType { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visGlueType, "GlueType"); } }
        public static SRC WalkPreference { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visWalkPref, "WalkPreference"); } }

        // GroupProperties
        public static SRC DisplayMode { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDisplayMode, "DisplayMode"); } }
        public static SRC DontMoveChildren { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDontMoveChildren, "DontMoveChildren"); } }
        public static SRC IsDropTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsDropTarget, "IsDropTarget"); } }
        public static SRC IsSnapTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsSnapTarget, "IsSnapTarget"); } }
        public static SRC IsTextEditTarget { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsTextEditTarget, "IsTextEditTarget"); } }
        public static SRC SelectMode { get { return new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupSelectMode, "SelectMode"); } }

        // Hyperlinks
        public static SRC Hyperlink_Address { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkAddress, "Address"); } }
        public static SRC Hyperlink_Default { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDefault, "Default"); } }
        public static SRC Hyperlink_Description { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDescription, "Description"); } }
        public static SRC Hyperlink_ExtraInfo { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkExtraInfo, "ExtraInfo"); } }
        public static SRC Hyperlink_Frame { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkFrame, "Frame"); } }
        public static SRC Hyperlink_Invisible { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkInvisible, "Invisible"); } }
        public static SRC Hyperlink_NewWindow { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkNewWin, "NewWindow"); } }
        public static SRC Hyperlink_SortKey { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSortKey, "SortKey"); } }
        public static SRC Hyperlink_SubAddress { get { return new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSubAddress, "SubAddress"); } }

        // Image Properties
        public static SRC Blur { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBlur, "Blur"); } }
        public static SRC Brightness { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBrightness, "Brightness"); } }
        public static SRC Contrast { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageContrast, "Contrast"); } }
        public static SRC Denoise { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageDenoise, "Denoise"); } }
        public static SRC Gamma { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageGamma, "Gamma"); } }
        public static SRC Sharpen { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageSharpen, "Sharpen"); } }
        public static SRC Transparency { get { return new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageTransparency, "Transparency"); } }

        // Line format
        public static SRC BeginArrow { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrow, "BeginArrow"); } }
        public static SRC BeginArrowSize { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrowSize, "BeginArrowSize"); } }
        public static SRC EndArrow { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrow, "EndArrow"); } }
        public static SRC EndArrowSize { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrowSize, "EndArrowSize"); } }
        public static SRC LineCap { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndCap, "LineCap"); } }
        public static SRC LineColor { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColor, "LineColor"); } }
        public static SRC LineColorTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColorTrans, "LineColorTrans"); } }
        public static SRC LinePattern { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLinePattern, "LinePattern"); } }
        public static SRC LineWeight { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineWeight, "LineWeight"); } }
        public static SRC Rounding { get { return new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineRounding, "Rounding"); } }

        // Miscellaneous
        public static SRC Calendar { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjCalendar, "Calendar"); } }
        public static SRC Comment { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visComment, "Comment"); } }
        public static SRC DropOnPageScale { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjDropOnPageScale, "DropOnPageScale"); } }
        public static SRC DynFeedback { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDynFeedback, "DynFeedback"); } }
        public static SRC IsDropSource { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDropSource, "IsDropSource"); } }
        public static SRC LangID { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLangID, "LangID"); } }
        public static SRC LocalizeMerge { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLocalizeMerge, "LocalizeMerge"); } }
        public static SRC NoAlignBox { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoAlignBox, "NoAlignBox"); } }
        public static SRC NoCtlHandles { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoCtlHandles, "NoCtlHandles"); } }
        public static SRC NoLiveDynamics { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoLiveDynamics, "NoLiveDynamics"); } }
        public static SRC NonPrinting { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNonPrinting, "NonPrinting"); } }
        public static SRC NoObjHandles { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoObjHandles, "NoObjHandles"); } }
        public static SRC ObjType { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visLOFlags, "ObjType"); } }
        public static SRC UpdateAlignBox { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visUpdateAlignBox, "UpdateAlignBox"); } }
        public static SRC HideText { get { return new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visHideText, "HideText"); } }

        // 1d endpoints
        public static SRC BeginX { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginX, "BeginX"); } }
        public static SRC BeginY { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginY, "BeginY"); } }
        public static SRC EndX { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndX, "EndX"); } }
        public static SRC EndY { get { return new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndY, "EndY"); } }

        // page layout
        public static SRC AvenueSizeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeX, "AvenueSizeX"); } }
        public static SRC AvenueSizeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeY, "AvenueSizeY"); } }
        public static SRC BlockSizeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeX, "BlockSizeX"); } }
        public static SRC BlockSizeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeY, "BlockSizeY"); } }
        public static SRC CtrlAsInput { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOCtrlAsInput, "CtrlAsInput"); } }
        public static SRC DynamicsOff { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLODynamicsOff, "DynamicsOff"); } }
        public static SRC EnableGrid { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOEnableGrid, "EnableGrid"); } }
        public static SRC LineAdjustFrom { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustFrom, "LineAdjustFrom"); } }
        public static SRC LineAdjustTo { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustTo, "LineAdjustTo"); } }
        public static SRC LineJumpCode { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpCode, "LineJumpCode"); } }
        public static SRC LineJumpFactorX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorX, "LineJumpFactorX"); } }
        public static SRC LineJumpFactorY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorY, "LineJumpFactorY"); } }
        public static SRC LineJumpStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpStyle, "LineJumpStyle"); } }
        public static SRC LineRouteExt { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineRouteExt, "LineRouteExt"); } }
        public static SRC LineToLineX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineX, "LineToLineX"); } }
        public static SRC LineToLineY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineY, "LineToLineY"); } }
        public static SRC LineToNodeX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeX, "LineToNodeX"); } }
        public static SRC LineToNodeY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeY, "LineToNodeY"); } }
        public static SRC PageLineJumpDirX { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirX, "PageLineJumpDirX"); } }
        public static SRC PageLineJumpDirY { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirY, "PageLineJumpDirY"); } }
        public static SRC PageShapeSplit { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOSplit, "PageShapeSplit"); } }
        public static SRC PlaceDepth { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceDepth, "PlaceDepth"); } }
        public static SRC PlaceFlip { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceFlip, "PlaceFlip"); } }
        public static SRC PlaceStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceStyle, "PlaceStyle"); } }
        public static SRC PlowCode { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlowCode, "PlowCode"); } }
        public static SRC ResizePage { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOResizePage, "ResizePage"); } }
        public static SRC RouteStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLORouteStyle, "RouteStyle"); } }

        public static SRC AvoidPageBreaks { get { return new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvoidPageBreaks, "AvoidPageBreaks"); } } // new in Visio 2010

        // print properties
        public static SRC PageLeftMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesLeftMargin, "PageLeftMargin"); } }
        public static SRC CenterX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterX, "CenterX"); } }
        public static SRC CenterY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterY, "CenterY"); } }
        public static SRC OnPage { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesOnPage, "OnPage"); } }
        public static SRC PageBottomMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesBottomMargin, "PageBottomMargin"); } }
        public static SRC PageRightMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesRightMargin, "PageRightMargin"); } }
        public static SRC PagesX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesX, "PagesX"); } }
        public static SRC PagesY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesY, "PagesY"); } }
        public static SRC PageTopMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesTopMargin, "PageTopMargin"); } }
        public static SRC PaperKind { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperKind, "PaperKind"); } }
        public static SRC PrintGrid { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPrintGrid, "PrintGrid"); } }
        public static SRC PrintPageOrientation { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPageOrientation, "PrintPageOrientation"); } }
        public static SRC ScaleX { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleX, "ScaleX"); } }
        public static SRC ScaleY { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleY, "ScaleY"); } }
        public static SRC PaperSource { get { return new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperSource, "PaperSource"); } }

        // page properties
        public static SRC DrawingScale { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawingScale, "DrawingScale"); } }
        public static SRC DrawingScaleType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawScaleType, "DrawingScaleType"); } }
        public static SRC DrawingSizeType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawSizeType, "DrawingSizeType"); } }
        public static SRC InhibitSnap { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageInhibitSnap, "InhibitSnap"); } }
        public static SRC PageHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageHeight, "PageHeight"); } }
        public static SRC PageScale { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageScale, "PageScale"); } }
        public static SRC PageWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageWidth, "PageWidth"); } }
        public static SRC ShdwObliqueAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwObliqueAngle, "ShdwObliqueAngle"); } }
        public static SRC ShdwOffsetX { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetX, "ShdwOffsetX"); } }
        public static SRC ShdwOffsetY { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetY, "ShdwOffsetY"); } }
        public static SRC ShdwScaleFactor { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwScaleFactor, "ShdwScaleFactor"); } }
        public static SRC ShdwType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwType, "ShdwType"); } }
        public static SRC UIVisibility { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageUIVisibility, "UIVisibility"); } }
        public static SRC DrawingResizeType { get { return new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawResizeType, "DrawingResizeType"); } } // new in Visio 2010

        // paragraph
        public static SRC Para_Bullet { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletIndex, "Bullet"); } }
        public static SRC Para_BulletFont { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFont, "BulletFont"); } }
        public static SRC Para_BulletFontSize { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFontSize, "BulletFontSize"); } }
        public static SRC Para_BulletStr { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletString, "BulletStr"); } }
        public static SRC Para_Flags { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visFlags, "Flags"); } }
        public static SRC Para_HorzAlign { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visHorzAlign, "HorzAlign"); } }
        public static SRC Para_IndFirst { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentFirst, "IndFirst"); } }
        public static SRC Para_IndLeft { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentLeft, "IndLeft"); } }
        public static SRC Para_IndRight { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentRight, "IndRight"); } }
        public static SRC Para_LocalizeBulletFont { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visLocalizeBulletFont, "LocalizeBulletFont"); } }
        public static SRC Para_SpAfter { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceAfter, "SpAfter"); } }
        public static SRC Para_SpBefore { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceBefore, "SpBefore"); } }
        public static SRC Para_SpLine { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceLine, "SpLine"); } }
        public static SRC Para_TextPosAfterBullet { get { return new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visTextPosAfterBullet, "TextPosAfterBullet"); } }

        // protection
        public static SRC LockAspect { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockAspect, "LockAspect"); } }
        public static SRC LockBegin { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockBegin, "LockBegin"); } }
        public static SRC LockCalcWH { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCalcWH, "LockCalcWH"); } }
        public static SRC LockCrop { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCrop, "LockCrop"); } }
        public static SRC LockCustProp { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCustProp, "LockCustProp"); } }
        public static SRC LockDelete { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockDelete, "LockDelete"); } }
        public static SRC LockEnd { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockEnd, "LockEnd"); } }
        public static SRC LockFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFormat, "LockFormat"); } }
        public static SRC LockFromGroupFormat { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFromGroupFormat, "LockFromGroupFormat"); } }
        public static SRC LockGroup { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockGroup, "LockGroup"); } }
        public static SRC LockHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockHeight, "LockHeight"); } }
        public static SRC LockMoveX { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveX, "LockMoveX"); } }
        public static SRC LockMoveY { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveY, "LockMoveY"); } }
        public static SRC LockRotate { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockRotate, "LockRotate"); } }
        public static SRC LockSelect { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockSelect, "LockSelect"); } }
        public static SRC LockTextEdit { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockTextEdit, "LockTextEdit"); } }
        public static SRC LockThemeColors { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeColors, "LockThemeColors"); } }
        public static SRC LockThemeEffects { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeEffects, "LockThemeEffects"); } }
        public static SRC LockVtxEdit { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockVtxEdit, "LockVtxEdit"); } }
        public static SRC LockWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockWidth, "LockWidth"); } }

        // ruler and grid
        public static SRC XGridDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridDensity, "XGridDensity"); } }
        public static SRC XGridOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridOrigin, "XGridOrigin"); } }
        public static SRC XGridSpacing { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridSpacing, "XGridSpacing"); } }
        public static SRC XRulerDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerDensity, "XRulerDensity"); } }
        public static SRC XRulerOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerOrigin, "XRulerOrigin"); } }
        public static SRC YGridDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridDensity, "YGridDensity"); } }
        public static SRC YGridOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridOrigin, "YGridOrigin"); } }
        public static SRC YGridSpacing { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridSpacing, "YGridSpacing"); } }
        public static SRC YRulerDensity { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerDensity, "YRulerDensity"); } }
        public static SRC YRulerOrigin { get { return new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerOrigin, "YRulerOrigin"); } }

        // Shape Tranform
        public static SRC Angle { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormAngle, "Angle"); } }
        public static SRC FlipX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipX, "FlipX"); } }
        public static SRC FlipY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipY, "FlipY"); } }
        public static SRC Height { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormHeight, "Height"); } }
        public static SRC LocPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinX, "LocPinX"); } }
        public static SRC LocPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinY, "LocPinY"); } }
        public static SRC PinX { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinX, "PinX"); } }
        public static SRC PinY { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinY, "PinY"); } }
        public static SRC ResizeMode { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormResizeMode, "ResizeMode"); } }
        public static SRC Width { get { return new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormWidth, "Width"); } }

        // reviewer
        public static SRC Reviewer_Color { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerColor, "Color"); } }
        public static SRC Reviewer_Initials { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerInitials, "Initials"); } }
        public static SRC Reviewer_Name { get { return new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerName, "Name"); } }

        // shape data
        public static SRC Prop_SortKey { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsSortKey, "SortKey"); } }
        public static SRC Prop_Ask { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsAsk, "Ask"); } }
        public static SRC Prop_Calendar { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsCalendar, "Calendar"); } }
        public static SRC Prop_Format { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsFormat, "Format"); } }
        public static SRC Prop_Invisible { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsInvis, "Invisible"); } }
        public static SRC Prop_Label { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLabel, "Label"); } }
        public static SRC Prop_LangID { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLangID, "LangID"); } }
        public static SRC Prop_Prompt { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsPrompt, "Prompt"); } }
        public static SRC Prop_Type { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsType, "Type"); } }
        public static SRC Prop_Value { get { return new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsValue, "Value"); } }

        // Layers
        public static SRC Layers_Active { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerActive, "Active"); } }
        public static SRC Layers_Color { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColor, "Color"); } }
        public static SRC Layers_Glue { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerGlue, "Glue"); } }
        public static SRC Layers_Locked { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerLock, "Locked"); } }
        public static SRC Layers_Print { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visDocPreviewScope, "Print"); } }
        public static SRC Layers_Snap { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerSnap, "Snap"); } }
        public static SRC Layers_ColorTrans { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColorTrans, "ColorTrans"); } }
        public static SRC Layers_Visible { get { return new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerVisible, "Visible"); } }

        //text transform
        public static SRC TxtAngle { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormAngle, "TxtAngle"); } }
        public static SRC TxtHeight { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormHeight, "TxtHeight"); } }
        public static SRC TxtLocPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinX, "TxtLocPinX"); } }
        public static SRC TxtLocPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinY, "TxtLocPinY"); } }
        public static SRC TxtPinX { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinX, "TxtPinX"); } }
        public static SRC TxtPinY { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinY, "TxtPinY"); } }
        public static SRC TxtWidth { get { return new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormWidth, "TxtWidth"); } }

        // user defined cells
        public static SRC User_Prompt { get { return new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserPrompt, "Prompt"); } }
        public static SRC User_Value { get { return new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserValue, "Value"); } }

        // Fields
        public static SRC Fields_Calendar { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCalendar, "Calendar"); } }
        public static SRC Fields_Format { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldFormat, "Format"); } }
        public static SRC Fields_ObjectKind { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldObjectKind, "ObjectKind"); } }
        public static SRC Fields_Type { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldType, "Type"); } }
        public static SRC Fields_UICat { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICategory, "UICat"); } }
        public static SRC Fields_UICod { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICode, "UICod"); } }
        public static SRC Fields_UIFmt { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUIFormat, "UIFmt"); } }
        public static SRC Fields_Value { get { return new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCell, "Value"); } }

        // text block format
        public static SRC BottomMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBottomMargin, "BottomMargin"); } }
        public static SRC DefaultTabStop { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDefaultTabStop, "DefaultTabStop"); } }
        public static SRC LeftMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkLeftMargin, "LeftMargin"); } }
        public static SRC RightMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkRightMargin, "RightMargin"); } }
        public static SRC TextBkgnd { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgnd, "TextBkgnd"); } }
        public static SRC TextBkgndTrans { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgndTrans, "TextBkgndTrans"); } }
        public static SRC TextDirection { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDirection, "TextDirection"); } }
        public static SRC TopMargin { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkTopMargin, "TopMargin"); } }
        public static SRC VerticalAlign { get { return new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkVerticalAlign, "VerticalAlign"); } }

        // Action tags
        public static SRC SmartTags_ButtonFace { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagButtonFace, "ButtonFace"); } }
        public static SRC SmartTags_Description { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDescription, "Description"); } }
        public static SRC SmartTags_Disabled { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisabled, "Disabled"); } }
        public static SRC SmartTags_DisplayMode { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisplayMode, "DisplayMode"); } }
        public static SRC SmartTags_TagName { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagName, "TagName"); } }
        public static SRC SmartTags_X { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagX, "X"); } }
        public static SRC SmartTags_XJustify { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagXJustify, "XJustify"); } }
        public static SRC SmartTags_Y { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagY, "Y"); } }
        public static SRC SmartTags_YJustify { get { return new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagYJustify, "YJustify"); } }

        // style
        public static SRC EnableFillProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesFill, "EnableFillProps"); } }
        public static SRC EnableLineProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesLine, "EnableLineProps"); } }
        public static SRC EnableTextProps { get { return new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesText, "EnableTextProps"); } }

        //tabs
        public static SRC Tabs_Alignment { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabAlign, "Alignment"); } }
        public static SRC Tabs_Position { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabPos, "Position"); } }
        public static SRC Tabs_StopCount { get { return new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabStopCount, "StopCount"); } }

        // shape layout
        public static SRC ConFixedCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOConFixedCode, "ConFixedCode"); } }
        public static SRC ConLineJumpCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpCode, "ConLineJumpCode"); } }
        public static SRC ConLineJumpDirX { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirX, "ConLineJumpDirX"); } }
        public static SRC ConLineJumpDirY { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirY, "ConLineJumpDirY"); } }
        public static SRC ConLineJumpStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpStyle, "ConLineJumpStyle"); } }
        public static SRC ConLineRouteExt { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOLineRouteExt, "ConLineRouteExt"); } }
        public static SRC ShapeFixedCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOFixedCode, "ShapeFixedCode"); } }
        public static SRC ShapePermeablePlace { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermeablePlace, "ShapePermeablePlace"); } }
        public static SRC ShapePermeableX { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermX, "ShapePermeableX"); } }
        public static SRC ShapePermeableY { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermY, "ShapePermeableY"); } }
        public static SRC ShapePlaceFlip { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceFlip, "ShapePlaceFlip"); } }
        public static SRC ShapePlaceStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceStyle, "ShapePlaceStyle"); } }
        public static SRC ShapePlowCode { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlowCode, "ShapePlowCode"); } }
        public static SRC ShapeRouteStyle { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORouteStyle, "ShapeRouteStyle"); } }
        public static SRC ShapeSplit { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplit, "ShapeSplit"); } }
        public static SRC ShapeSplittable { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplittable, "ShapeSplittable"); } }
        public static SRC DisplayLevel { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLODisplayLevel, "DisplayLevel"); } } // new in Visio 2010
        public static SRC Relationships { get { return new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORelationships, "Relationships"); } } // new in Visio 2010




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