using System.Collections.Generic;
using System.Linq;
using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;

namespace VisioAutomation.ShapeSheet
{
    public static class SRCConstants
    {
        // Actions
        public static SRC Actions_Action => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionAction);
        public static SRC Actions_BeginGroup => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionBeginGroup);
        public static SRC Actions_ButtonFace => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionButtonFace);
        public static SRC Actions_Checked => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionChecked);
        public static SRC Actions_Disabled => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionDisabled);
        public static SRC Actions_Invisible => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionInvisible);
        public static SRC Actions_Menu => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionMenu);
        public static SRC Actions_ReadOnly => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionReadOnly);
        public static SRC Actions_SortKey => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionSortKey);
        public static SRC Actions_TagName => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionTagName);
        public static SRC Actions_FlyoutChild => new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionFlyoutChild);
// new for visio 2010

        // Alignment
        public static SRC AlignBottom => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignBottom);
        public static SRC AlignCenter => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignCenter);
        public static SRC AlignLeft => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignLeft);
        public static SRC AlignMiddle => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignMiddle);
        public static SRC AlignRight => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignRight);
        public static SRC AlignTop => new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignTop);

        // Annotation
        public static SRC Annotation_Comment => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationComment);
        public static SRC Annotation_Date => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationDate);
        public static SRC Annotation_LangID => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationLangID);
        public static SRC Annotation_MarkerIndex => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationMarkerIndex);
        public static SRC Annotation_X => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationX);
        public static SRC Annotation_Y => new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationY);

        // Character
        public static SRC CharAsianFont => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterAsianFont);
        public static SRC CharCase => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterCase);
        public static SRC CharColor => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColor);
        public static SRC CharComplexScriptFont => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptFont);
        public static SRC CharComplexScriptSize => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptSize);
        public static SRC CharDoubleStrikethrough => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDoubleStrikethrough);
        public static SRC CharDblUnderline => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDblUnderline);
        public static SRC CharFont => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFont);
        public static SRC CharLangID => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLangID);
        public static SRC CharLocale => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocale);
        public static SRC CharLocalizeFont => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocalizeFont);
        public static SRC CharOverline => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterOverline);
        public static SRC CharPerpendicular => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPerpendicular);
        public static SRC CharPos => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPos);
        public static SRC CharRTLText => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterRTLText);
        public static SRC CharFontScale => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFontScale);
        public static SRC CharSize => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterSize);
        public static SRC CharLetterspace => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLetterspace);
        public static SRC CharStrikethru => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStrikethru);
        public static SRC CharStyle => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStyle);
        public static SRC CharColorTrans => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColorTrans);
        public static SRC CharUseVertical => new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterUseVertical);

        // Connections
        public static SRC Connections_D => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctD);
        public static SRC Connections_DirX => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirX);
        public static SRC Connections_DirY => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirY);
        public static SRC Connections_Type => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctType);
        public static SRC Connections_X => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visX);
        public static SRC Connections_Y => new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visY);

        // Controls
        public static SRC Controls_CanGlue => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlGlue);
        public static SRC Controls_Tip => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlTip);
        public static SRC Controls_XCon => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXCon);
        public static SRC Controls_X => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlX);
        public static SRC Controls_XDyn => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXDyn);
        public static SRC Controls_YCon => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYCon);
        public static SRC Controls_Y => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlY);
        public static SRC Controls_YDyn => new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYDyn);

        // Document Properties
        public static SRC AddMarkup => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocAddMarkup);
        public static SRC DocLangID => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLangID);
        public static SRC LockPreview => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLockPreview);
        public static SRC OutputFormat => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocOutputFormat);
        public static SRC PreviewQuality => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewQuality);
        public static SRC PreviewScope => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewScope);
        public static SRC ViewMarkup => new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocViewMarkup);

        // Events
        public static SRC EventDblClick => new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDblClick);
        public static SRC EventDrop => new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDrop);
        public static SRC EventMultiDrop => new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellMultiDrop);
        public static SRC EventXFMod => new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellXFMod);
        public static SRC TheText => new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellTheText);

        // ForeignImageInfo
        public static SRC ImgHeight => new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgHeight);
        public static SRC ImgOffsetX => new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetX);
        public static SRC ImgOffsetY => new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetY);
        public static SRC ImgWidth => new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgWidth);

        // Geometry
        public static SRC Geometry_A => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visBow);
        public static SRC Geometry_B => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visControlX);
        public static SRC Geometry_C => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visEccentricityAngle);
        public static SRC Geometry_D => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visAspectRatio);
        public static SRC Geometry_E => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visNURBSData);
        public static SRC Geometry_X => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visX);
        public static SRC Geometry_Y => new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visY);
        public static SRC Geometry_NoFill => new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoFill);
        public static SRC Geometry_NoLine => new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoLine);
        public static SRC Geometry_NoShow => new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoShow);
        public static SRC Geometry_NoSnap => new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoSnap);
        public static SRC Geometry_NoQuickDrag => new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoQuickDrag);

        // Fill Format
        public static SRC FillBkgnd => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgnd);
        public static SRC FillBkgndTrans => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgndTrans);
        public static SRC FillForegnd => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd);
        public static SRC FillForegndTrans => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegndTrans);
        public static SRC FillPattern => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillPattern);
        public static SRC ShapeShdwObliqueAngle => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwObliqueAngle);
        public static SRC ShapeShdwOffsetX => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetX);
        public static SRC ShapeShdwOffsetY => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetY);
        public static SRC ShapeShdwScaleFactor => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwScaleFactor);
        public static SRC ShapeShdwType => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwType);
        public static SRC ShdwBkgnd => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgnd);
        public static SRC ShdwBkgndTrans => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgndTrans);
        public static SRC ShdwForegnd => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegnd);
        public static SRC ShdwForegndTrans => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegndTrans);
        public static SRC ShdwPattern => new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwPattern);

        // GlueInfo
        public static SRC BegTrigger => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visBegTrigger);
        public static SRC EndTrigger => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visEndTrigger);
        public static SRC GlueType => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visGlueType);
        public static SRC WalkPreference => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visWalkPref);

        // GroupProperties
        public static SRC DisplayMode => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDisplayMode);
        public static SRC DontMoveChildren => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDontMoveChildren);
        public static SRC IsDropTarget => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsDropTarget);
        public static SRC IsSnapTarget => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsSnapTarget);
        public static SRC IsTextEditTarget => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsTextEditTarget);
        public static SRC SelectMode => new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupSelectMode);

        // Hyperlinks
        public static SRC Hyperlink_Address => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkAddress);
        public static SRC Hyperlink_Default => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDefault);
        public static SRC Hyperlink_Description => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDescription);
        public static SRC Hyperlink_ExtraInfo => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkExtraInfo);
        public static SRC Hyperlink_Frame => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkFrame);
        public static SRC Hyperlink_Invisible => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkInvisible);
        public static SRC Hyperlink_NewWindow => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkNewWin);
        public static SRC Hyperlink_SortKey => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSortKey);
        public static SRC Hyperlink_SubAddress => new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSubAddress);

        // Image Properties
        public static SRC Blur => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBlur);
        public static SRC Brightness => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBrightness);
        public static SRC Contrast => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageContrast);
        public static SRC Denoise => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageDenoise);
        public static SRC Gamma => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageGamma);
        public static SRC Sharpen => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageSharpen);
        public static SRC Transparency => new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageTransparency);

        // Line format
        public static SRC BeginArrow => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrow);
        public static SRC BeginArrowSize => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrowSize);
        public static SRC EndArrow => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrow);
        public static SRC EndArrowSize => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrowSize);
        public static SRC LineCap => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndCap);
        public static SRC LineColor => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColor);
        public static SRC LineColorTrans => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColorTrans);
        public static SRC LinePattern => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLinePattern);
        public static SRC LineWeight => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineWeight);
        public static SRC Rounding => new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineRounding);

        // Miscellaneous
        public static SRC Calendar => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjCalendar);
        public static SRC Comment => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visComment);
        public static SRC DropOnPageScale => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjDropOnPageScale);
        public static SRC DynFeedback => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDynFeedback);
        public static SRC IsDropSource => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDropSource);
        public static SRC LangID => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLangID);
        public static SRC LocalizeMerge => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLocalizeMerge);
        public static SRC NoAlignBox => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoAlignBox);
        public static SRC NoCtlHandles => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoCtlHandles);
        public static SRC NoLiveDynamics => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoLiveDynamics);
        public static SRC NonPrinting => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNonPrinting);
        public static SRC NoObjHandles => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoObjHandles);
        public static SRC ObjType => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visLOFlags);
        public static SRC UpdateAlignBox => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visUpdateAlignBox);
        public static SRC HideText => new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visHideText);

        // 1d endpoints
        public static SRC BeginX => new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginX);
        public static SRC BeginY => new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginY);
        public static SRC EndX => new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndX);
        public static SRC EndY => new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndY);

        // page layout
        public static SRC AvenueSizeX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeX);
        public static SRC AvenueSizeY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeY);
        public static SRC BlockSizeX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeX);
        public static SRC BlockSizeY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeY);
        public static SRC CtrlAsInput => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOCtrlAsInput);
        public static SRC DynamicsOff => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLODynamicsOff);
        public static SRC EnableGrid => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOEnableGrid);
        public static SRC LineAdjustFrom => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustFrom);
        public static SRC LineAdjustTo => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustTo);
        public static SRC LineJumpCode => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpCode);
        public static SRC LineJumpFactorX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorX);
        public static SRC LineJumpFactorY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorY);
        public static SRC LineJumpStyle => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpStyle);
        public static SRC LineRouteExt => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineRouteExt);
        public static SRC LineToLineX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineX);
        public static SRC LineToLineY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineY);
        public static SRC LineToNodeX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeX);
        public static SRC LineToNodeY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeY);
        public static SRC PageLineJumpDirX => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirX);
        public static SRC PageLineJumpDirY => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirY);
        public static SRC PageShapeSplit => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOSplit);
        public static SRC PlaceDepth => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceDepth);
        public static SRC PlaceFlip => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceFlip);
        public static SRC PlaceStyle => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceStyle);
        public static SRC PlowCode => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlowCode);
        public static SRC ResizePage => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOResizePage);
        public static SRC RouteStyle => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLORouteStyle);

        public static SRC AvoidPageBreaks => new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvoidPageBreaks);
// new in Visio 2010

        // print properties
        public static SRC PageLeftMargin => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesLeftMargin);
        public static SRC CenterX => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterX);
        public static SRC CenterY => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterY);
        public static SRC OnPage => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesOnPage);
        public static SRC PageBottomMargin => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesBottomMargin);
        public static SRC PageRightMargin => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesRightMargin);
        public static SRC PagesX => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesX);
        public static SRC PagesY => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesY);
        public static SRC PageTopMargin => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesTopMargin);
        public static SRC PaperKind => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperKind);
        public static SRC PrintGrid => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPrintGrid);
        public static SRC PrintPageOrientation => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPageOrientation);
        public static SRC ScaleX => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleX);
        public static SRC ScaleY => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleY);
        public static SRC PaperSource => new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperSource);

        // page properties
        public static SRC DrawingScale => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawingScale);
        public static SRC DrawingScaleType => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawScaleType);
        public static SRC DrawingSizeType => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawSizeType);
        public static SRC InhibitSnap => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageInhibitSnap);
        public static SRC PageHeight => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageHeight);
        public static SRC PageScale => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageScale);
        public static SRC PageWidth => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageWidth);
        public static SRC ShdwObliqueAngle => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwObliqueAngle);
        public static SRC ShdwOffsetX => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetX);
        public static SRC ShdwOffsetY => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetY);
        public static SRC ShdwScaleFactor => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwScaleFactor);
        public static SRC ShdwType => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwType);
        public static SRC UIVisibility => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageUIVisibility);
        public static SRC DrawingResizeType => new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawResizeType);
// new in Visio 2010

        // paragraph
        public static SRC Para_Bullet => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletIndex);
        public static SRC Para_BulletFont => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFont);
        public static SRC Para_BulletFontSize => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFontSize);
        public static SRC Para_BulletStr => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletString);
        public static SRC Para_Flags => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visFlags);
        public static SRC Para_HorzAlign => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visHorzAlign);
        public static SRC Para_IndFirst => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentFirst);
        public static SRC Para_IndLeft => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentLeft);
        public static SRC Para_IndRight => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentRight);
        public static SRC Para_LocalizeBulletFont => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visLocalizeBulletFont);
        public static SRC Para_SpAfter => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceAfter);
        public static SRC Para_SpBefore => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceBefore);
        public static SRC Para_SpLine => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceLine);
        public static SRC Para_TextPosAfterBullet => new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visTextPosAfterBullet);

        // protection
        public static SRC LockAspect => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockAspect);
        public static SRC LockBegin => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockBegin);
        public static SRC LockCalcWH => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCalcWH);
        public static SRC LockCrop => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCrop);
        public static SRC LockCustProp => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCustProp);
        public static SRC LockDelete => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockDelete);
        public static SRC LockEnd => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockEnd);
        public static SRC LockFormat => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFormat);
        public static SRC LockFromGroupFormat => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFromGroupFormat);
        public static SRC LockGroup => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockGroup);
        public static SRC LockHeight => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockHeight);
        public static SRC LockMoveX => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveX);
        public static SRC LockMoveY => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveY);
        public static SRC LockRotate => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockRotate);
        public static SRC LockSelect => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockSelect);
        public static SRC LockTextEdit => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockTextEdit);
        public static SRC LockThemeColors => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeColors);
        public static SRC LockThemeEffects => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeEffects);
        public static SRC LockVtxEdit => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockVtxEdit);
        public static SRC LockWidth => new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockWidth);

        // ruler and grid
        public static SRC XGridDensity => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridDensity);
        public static SRC XGridOrigin => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridOrigin);
        public static SRC XGridSpacing => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridSpacing);
        public static SRC XRulerDensity => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerDensity);
        public static SRC XRulerOrigin => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerOrigin);
        public static SRC YGridDensity => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridDensity);
        public static SRC YGridOrigin => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridOrigin);
        public static SRC YGridSpacing => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridSpacing);
        public static SRC YRulerDensity => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerDensity);
        public static SRC YRulerOrigin => new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerOrigin);

        // Shape Tranform
        public static SRC Angle => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormAngle);
        public static SRC FlipX => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipX);
        public static SRC FlipY => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipY);
        public static SRC Height => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormHeight);
        public static SRC LocPinX => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinX);
        public static SRC LocPinY => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinY);
        public static SRC PinX => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinX);
        public static SRC PinY => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinY);
        public static SRC ResizeMode => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormResizeMode);
        public static SRC Width => new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormWidth);

        // reviewer
        public static SRC Reviewer_Color => new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerColor);
        public static SRC Reviewer_Initials => new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerInitials);
        public static SRC Reviewer_Name => new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerName);

        // shape data
        public static SRC Prop_SortKey => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsSortKey);
        public static SRC Prop_Ask => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsAsk);
        public static SRC Prop_Calendar => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsCalendar);
        public static SRC Prop_Format => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsFormat);
        public static SRC Prop_Invisible => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsInvis);
        public static SRC Prop_Label => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLabel);
        public static SRC Prop_LangID => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLangID);
        public static SRC Prop_Prompt => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsPrompt);
        public static SRC Prop_Type => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsType);
        public static SRC Prop_Value => new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsValue);

        // Layers
        public static SRC Layers_Active => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerActive);
        public static SRC Layers_Color => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColor);
        public static SRC Layers_Glue => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerGlue);
        public static SRC Layers_Locked => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerLock);
        public static SRC Layers_Print => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visDocPreviewScope);
        public static SRC Layers_Snap => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerSnap);
        public static SRC Layers_ColorTrans => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColorTrans);
        public static SRC Layers_Visible => new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerVisible);

        //text transform
        public static SRC TxtAngle => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormAngle);
        public static SRC TxtHeight => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormHeight);
        public static SRC TxtLocPinX => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinX);
        public static SRC TxtLocPinY => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinY);
        public static SRC TxtPinX => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinX);
        public static SRC TxtPinY => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinY);
        public static SRC TxtWidth => new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormWidth);

        // user defined cells
        public static SRC User_Prompt => new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserPrompt);
        public static SRC User_Value => new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserValue);

        // Fields
        public static SRC Fields_Calendar => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCalendar);
        public static SRC Fields_Format => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldFormat);
        public static SRC Fields_ObjectKind => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldObjectKind);
        public static SRC Fields_Type => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldType);
        public static SRC Fields_UICat => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICategory);
        public static SRC Fields_UICod => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICode);
        public static SRC Fields_UIFmt => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUIFormat);
        public static SRC Fields_Value => new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCell);

        // text block format
        public static SRC BottomMargin => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBottomMargin);
        public static SRC DefaultTabStop => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDefaultTabStop);
        public static SRC LeftMargin => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkLeftMargin);
        public static SRC RightMargin => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkRightMargin);
        public static SRC TextBkgnd => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgnd);
        public static SRC TextBkgndTrans => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgndTrans);
        public static SRC TextDirection => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDirection);
        public static SRC TopMargin => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkTopMargin);
        public static SRC VerticalAlign => new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkVerticalAlign);

        // Action tags
        public static SRC SmartTags_ButtonFace => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagButtonFace);
        public static SRC SmartTags_Description => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDescription);
        public static SRC SmartTags_Disabled => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisabled);
        public static SRC SmartTags_DisplayMode => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisplayMode);
        public static SRC SmartTags_TagName => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagName);
        public static SRC SmartTags_X => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagX);
        public static SRC SmartTags_XJustify => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagXJustify);
        public static SRC SmartTags_Y => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagY);
        public static SRC SmartTags_YJustify => new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagYJustify);

        // style
        public static SRC EnableFillProps => new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesFill);
        public static SRC EnableLineProps => new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesLine);
        public static SRC EnableTextProps => new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesText);

        //tabs
        public static SRC Tabs_Alignment => new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabAlign);
        public static SRC Tabs_Position => new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabPos);
        public static SRC Tabs_StopCount => new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabStopCount);

        // shape layout
        public static SRC ConFixedCode => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOConFixedCode);
        public static SRC ConLineJumpCode => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpCode);
        public static SRC ConLineJumpDirX => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirX);
        public static SRC ConLineJumpDirY => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirY);
        public static SRC ConLineJumpStyle => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpStyle);
        public static SRC ConLineRouteExt => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOLineRouteExt);
        public static SRC ShapeFixedCode => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOFixedCode);
        public static SRC ShapePermeablePlace => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermeablePlace);
        public static SRC ShapePermeableX => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermX);
        public static SRC ShapePermeableY => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermY);
        public static SRC ShapePlaceFlip => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceFlip);
        public static SRC ShapePlaceStyle => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceStyle);
        public static SRC ShapePlowCode => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlowCode);
        public static SRC ShapeRouteStyle => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORouteStyle);
        public static SRC ShapeSplit => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplit);
        public static SRC ShapeSplittable => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplittable);
        public static SRC DisplayLevel => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLODisplayLevel);
// new in Visio 2010
        public static SRC Relationships => new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORelationships);
// new in Visio 2010

        public static Dictionary<string, SRC> GetSRCDictionary()
        {
            var srcconstants_t = typeof(SRCConstants);

            var binding_flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static;

            // find all the static properties that return SRC types
            var src_type = typeof(SRC);
            var props = srcconstants_t.GetProperties(binding_flags)
                .Where(p => p.PropertyType == src_type);

            var fields_name_to_value = new Dictionary<string, SRC>();
            foreach (var propinfo in props)
            {
                var src = (SRC)propinfo.GetValue(null, null);
                var name = propinfo.Name;
                fields_name_to_value[name] = src;
            }

            return fields_name_to_value;
        }
    }
}
