using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;

namespace VisioAutomation.ShapeSheet
{
    public static class SRCConstants
    {
        // Actions
        public static Src Actions_Action => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionAction);
        public static Src Actions_BeginGroup => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionBeginGroup);
        public static Src Actions_ButtonFace => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionButtonFace);
        public static Src Actions_Checked => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionChecked);
        public static Src Actions_Disabled => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionDisabled);
        public static Src Actions_Invisible => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionInvisible);
        public static Src Actions_Menu => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionMenu);
        public static Src Actions_ReadOnly => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionReadOnly);
        public static Src Actions_SortKey => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionSortKey);
        public static Src Actions_TagName => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionTagName);
        public static Src Actions_FlyoutChild => new Src(SEC.visSectionAction, ROW.visRowAction, CEL.visActionFlyoutChild); // new for visio 2010

        // Alignment
        public static Src AlignBottom => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignBottom);
        public static Src AlignCenter => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignCenter);
        public static Src AlignLeft => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignLeft);
        public static Src AlignMiddle => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignMiddle);
        public static Src AlignRight => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignRight);
        public static Src AlignTop => new Src(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignTop);

        // Annotation
        public static Src Annotation_Comment => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationComment);
        public static Src Annotation_Date => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationDate);
        public static Src Annotation_LangID => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationLangID);
        public static Src Annotation_MarkerIndex => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationMarkerIndex);
        public static Src Annotation_X => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationX);
        public static Src Annotation_Y => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationY);

        // Character
        public static Src CharAsianFont => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterAsianFont);
        public static Src CharCase => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterCase);
        public static Src CharColor => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColor);
        public static Src CharComplexScriptFont => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptFont);
        public static Src CharComplexScriptSize => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptSize);
        public static Src CharDoubleStrikethrough => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDoubleStrikethrough);
        public static Src CharDblUnderline => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDblUnderline);
        public static Src CharFont => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFont);
        public static Src CharLangID => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLangID);
        public static Src CharLocale => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocale);
        public static Src CharLocalizeFont => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocalizeFont);
        public static Src CharOverline => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterOverline);
        public static Src CharPerpendicular => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPerpendicular);
        public static Src CharPos => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPos);
        public static Src CharRTLText => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterRTLText);
        public static Src CharFontScale => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFontScale);
        public static Src CharSize => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterSize);
        public static Src CharLetterspace => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLetterspace);
        public static Src CharStrikethru => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStrikethru);
        public static Src CharStyle => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStyle);
        public static Src CharColorTrans => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColorTrans);
        public static Src CharUseVertical => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterUseVertical);

        // Connections
        public static Src Connections_D => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctD);
        public static Src Connections_DirX => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirX);
        public static Src Connections_DirY => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirY);
        public static Src Connections_Type => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctType);
        public static Src Connections_X => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visX);
        public static Src Connections_Y => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visY);

        // Controls
        public static Src Controls_CanGlue => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlGlue);
        public static Src Controls_Tip => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlTip);
        public static Src Controls_XCon => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXCon);
        public static Src Controls_X => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlX);
        public static Src Controls_XDyn => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXDyn);
        public static Src Controls_YCon => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYCon);
        public static Src Controls_Y => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlY);
        public static Src Controls_YDyn => new Src(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYDyn);

        // Document Properties
        public static Src AddMarkup => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocAddMarkup);
        public static Src DocLangID => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLangID);
        public static Src LockPreview => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLockPreview);
        public static Src OutputFormat => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocOutputFormat);
        public static Src PreviewQuality => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewQuality);
        public static Src PreviewScope => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewScope);
        public static Src ViewMarkup => new Src(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocViewMarkup);

        // Events
        public static Src EventDblClick => new Src(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDblClick);
        public static Src EventDrop => new Src(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDrop);
        public static Src EventMultiDrop => new Src(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellMultiDrop);
        public static Src EventXFMod => new Src(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellXFMod);
        public static Src TheText => new Src(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellTheText);

        // ForeignImageInfo
        public static Src ImgHeight => new Src(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgHeight);
        public static Src ImgOffsetX => new Src(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetX);
        public static Src ImgOffsetY => new Src(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetY);
        public static Src ImgWidth => new Src(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgWidth);

        // Geometry
        public static Src Geometry_A => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visBow);
        public static Src Geometry_B => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visControlX);
        public static Src Geometry_C => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visEccentricityAngle);
        public static Src Geometry_D => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visAspectRatio);
        public static Src Geometry_E => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visNURBSData);
        public static Src Geometry_X => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visX);
        public static Src Geometry_Y => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visY);
        public static Src Geometry_NoFill => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoFill);
        public static Src Geometry_NoLine => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoLine);
        public static Src Geometry_NoShow => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoShow);
        public static Src Geometry_NoSnap => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoSnap);
        public static Src Geometry_NoQuickDrag => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoQuickDrag);

        // Fill Format
        public static Src FillBkgnd => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgnd);
        public static Src FillBkgndTrans => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgndTrans);
        public static Src FillForegnd => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd);
        public static Src FillForegndTrans => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegndTrans);
        public static Src FillPattern => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillPattern);
        public static Src ShapeShdwObliqueAngle => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwObliqueAngle);
        public static Src ShapeShdwOffsetX => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetX);
        public static Src ShapeShdwOffsetY => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetY);
        public static Src ShapeShdwScaleFactor => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwScaleFactor);
        public static Src ShapeShdwType => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwType);
        public static Src ShdwBkgnd => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgnd);
        public static Src ShdwBkgndTrans => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgndTrans);
        public static Src ShdwForegnd => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegnd);
        public static Src ShdwForegndTrans => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegndTrans);
        public static Src ShdwPattern => new Src(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwPattern);

        // GlueInfo
        public static Src BegTrigger => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visBegTrigger);
        public static Src EndTrigger => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visEndTrigger);
        public static Src GlueType => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visGlueType);
        public static Src WalkPreference => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visWalkPref);

        // GroupProperties
        public static Src DisplayMode => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDisplayMode);
        public static Src DontMoveChildren => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDontMoveChildren);
        public static Src IsDropTarget => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsDropTarget);
        public static Src IsSnapTarget => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsSnapTarget);
        public static Src IsTextEditTarget => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsTextEditTarget);
        public static Src SelectMode => new Src(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupSelectMode);

        // Hyperlinks
        public static Src Hyperlink_Address => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkAddress);
        public static Src Hyperlink_Default => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkDefault);
        public static Src Hyperlink_Description => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkDescription);
        public static Src Hyperlink_ExtraInfo => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkExtraInfo);
        public static Src Hyperlink_Frame => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkFrame);
        public static Src Hyperlink_Invisible => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkInvisible);
        public static Src Hyperlink_NewWindow => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkNewWin);
        public static Src Hyperlink_SortKey => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkSortKey);
        public static Src Hyperlink_SubAddress => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, CEL.visHLinkSubAddress);

        // Image Properties
        public static Src Blur => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBlur);
        public static Src Brightness => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBrightness);
        public static Src Contrast => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageContrast);
        public static Src Denoise => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageDenoise);
        public static Src Gamma => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageGamma);
        public static Src Sharpen => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageSharpen);
        public static Src Transparency => new Src(SEC.visSectionObject, ROW.visRowImage, CEL.visImageTransparency);

        // Line format
        public static Src BeginArrow => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrow);
        public static Src BeginArrowSize => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrowSize);
        public static Src EndArrow => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrow);
        public static Src EndArrowSize => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrowSize);
        public static Src LineCap => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndCap);
        public static Src LineColor => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColor);
        public static Src LineColorTrans => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColorTrans);
        public static Src LinePattern => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLinePattern);
        public static Src LineWeight => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineWeight);
        public static Src Rounding => new Src(SEC.visSectionObject, ROW.visRowLine, CEL.visLineRounding);

        // Miscellaneous
        public static Src Calendar => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjCalendar);
        public static Src Comment => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visComment);
        public static Src DropOnPageScale => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjDropOnPageScale);
        public static Src DynFeedback => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visDynFeedback);
        public static Src IsDropSource => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visDropSource);
        public static Src LangID => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLangID);
        public static Src LocalizeMerge => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLocalizeMerge);
        public static Src NoAlignBox => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoAlignBox);
        public static Src NoCtlHandles => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoCtlHandles);
        public static Src NoLiveDynamics => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoLiveDynamics);
        public static Src NonPrinting => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visNonPrinting);
        public static Src NoObjHandles => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoObjHandles);
        public static Src ObjType => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visLOFlags);
        public static Src UpdateAlignBox => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visUpdateAlignBox);
        public static Src HideText => new Src(SEC.visSectionObject, ROW.visRowMisc, CEL.visHideText);

        // 1d endpoints
        public static Src BeginX => new Src(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginX);
        public static Src BeginY => new Src(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginY);
        public static Src EndX => new Src(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndX);
        public static Src EndY => new Src(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndY);

        // page layout
        public static Src AvenueSizeX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeX);
        public static Src AvenueSizeY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeY);
        public static Src BlockSizeX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeX);
        public static Src BlockSizeY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeY);
        public static Src CtrlAsInput => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOCtrlAsInput);
        public static Src DynamicsOff => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLODynamicsOff);
        public static Src EnableGrid => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOEnableGrid);
        public static Src LineAdjustFrom => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustFrom);
        public static Src LineAdjustTo => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustTo);
        public static Src LineJumpCode => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpCode);
        public static Src LineJumpFactorX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorX);
        public static Src LineJumpFactorY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorY);
        public static Src LineJumpStyle => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpStyle);
        public static Src LineRouteExt => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineRouteExt);
        public static Src LineToLineX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineX);
        public static Src LineToLineY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineY);
        public static Src LineToNodeX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeX);
        public static Src LineToNodeY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeY);
        public static Src PageLineJumpDirX => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirX);
        public static Src PageLineJumpDirY => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirY);
        public static Src PageShapeSplit => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOSplit);
        public static Src PlaceDepth => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceDepth);
        public static Src PlaceFlip => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceFlip);
        public static Src PlaceStyle => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceStyle);
        public static Src PlowCode => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlowCode);
        public static Src ResizePage => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOResizePage);
        public static Src RouteStyle => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLORouteStyle);

        public static Src AvoidPageBreaks => new Src(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvoidPageBreaks); // new in Visio 2010

        // print properties
        public static Src PageLeftMargin => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesLeftMargin);
        public static Src CenterX => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterX);
        public static Src CenterY => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterY);
        public static Src OnPage => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesOnPage);
        public static Src PageBottomMargin => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesBottomMargin);
        public static Src PageRightMargin => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesRightMargin);
        public static Src PagesX => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesX);
        public static Src PagesY => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesY);
        public static Src PageTopMargin => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesTopMargin);
        public static Src PaperKind => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperKind);
        public static Src PrintGrid => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPrintGrid);
        public static Src PrintPageOrientation => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPageOrientation);
        public static Src ScaleX => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleX);
        public static Src ScaleY => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleY);
        public static Src PaperSource => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperSource);

        // page properties
        public static Src DrawingScale => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawingScale);
        public static Src DrawingScaleType => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawScaleType);
        public static Src DrawingSizeType => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawSizeType);
        public static Src InhibitSnap => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageInhibitSnap);
        public static Src PageHeight => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageHeight);
        public static Src PageScale => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageScale);
        public static Src PageWidth => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageWidth);
        public static Src ShdwObliqueAngle => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwObliqueAngle);
        public static Src ShdwOffsetX => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetX);
        public static Src ShdwOffsetY => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetY);
        public static Src ShdwScaleFactor => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwScaleFactor);
        public static Src ShdwType => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwType);
        public static Src UIVisibility => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageUIVisibility);
        public static Src DrawingResizeType => new Src(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawResizeType); // new in Visio 2010

        // paragraph
        public static Src Para_Bullet => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletIndex);
        public static Src Para_BulletFont => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFont);
        public static Src Para_BulletFontSize => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFontSize);
        public static Src Para_BulletStr => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletString);
        public static Src Para_Flags => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visFlags);
        public static Src Para_HorzAlign => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visHorzAlign);
        public static Src Para_IndFirst => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentFirst);
        public static Src Para_IndLeft => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentLeft);
        public static Src Para_IndRight => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentRight);
        public static Src Para_LocalizeBulletFont => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visLocalizeBulletFont);
        public static Src Para_SpAfter => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceAfter);
        public static Src Para_SpBefore => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceBefore);
        public static Src Para_SpLine => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceLine);
        public static Src Para_TextPosAfterBullet => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visTextPosAfterBullet);

        // protection
        public static Src LockAspect => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockAspect);
        public static Src LockBegin => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockBegin);
        public static Src LockCalcWH => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCalcWH);
        public static Src LockCrop => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCrop);
        public static Src LockCustProp => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCustProp);
        public static Src LockDelete => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockDelete);
        public static Src LockEnd => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockEnd);
        public static Src LockFormat => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFormat);
        public static Src LockFromGroupFormat => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFromGroupFormat);
        public static Src LockGroup => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockGroup);
        public static Src LockHeight => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockHeight);
        public static Src LockMoveX => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveX);
        public static Src LockMoveY => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveY);
        public static Src LockRotate => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockRotate);
        public static Src LockSelect => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockSelect);
        public static Src LockTextEdit => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockTextEdit);
        public static Src LockThemeColors => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeColors);
        public static Src LockThemeEffects => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeEffects);
        public static Src LockVtxEdit => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockVtxEdit);
        public static Src LockWidth => new Src(SEC.visSectionObject, ROW.visRowLock, CEL.visLockWidth);

        // ruler and grid
        public static Src XGridDensity => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridDensity);
        public static Src XGridOrigin => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridOrigin);
        public static Src XGridSpacing => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridSpacing);
        public static Src XRulerDensity => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerDensity);
        public static Src XRulerOrigin => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerOrigin);
        public static Src YGridDensity => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridDensity);
        public static Src YGridOrigin => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridOrigin);
        public static Src YGridSpacing => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridSpacing);
        public static Src YRulerDensity => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerDensity);
        public static Src YRulerOrigin => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerOrigin);

        // Shape Tranform
        public static Src Angle => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormAngle);
        public static Src FlipX => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipX);
        public static Src FlipY => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipY);
        public static Src Height => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormHeight);
        public static Src LocPinX => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinX);
        public static Src LocPinY => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinY);
        public static Src PinX => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinX);
        public static Src PinY => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinY);
        public static Src ResizeMode => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormResizeMode);
        public static Src Width => new Src(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormWidth);

        // reviewer
        public static Src Reviewer_Color => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerColor);
        public static Src Reviewer_Initials => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerInitials);
        public static Src Reviewer_Name => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerName);

        // shape data
        public static Src Prop_SortKey => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsSortKey);
        public static Src Prop_Ask => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsAsk);
        public static Src Prop_Calendar => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsCalendar);
        public static Src Prop_Format => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsFormat);
        public static Src Prop_Invisible => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsInvis);
        public static Src Prop_Label => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLabel);
        public static Src Prop_LangID => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLangID);
        public static Src Prop_Prompt => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsPrompt);
        public static Src Prop_Type => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsType);
        public static Src Prop_Value => new Src(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsValue);

        // Layers
        public static Src Layers_Active => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerActive);
        public static Src Layers_Color => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColor);
        public static Src Layers_Glue => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerGlue);
        public static Src Layers_Locked => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerLock);
        public static Src Layers_Print => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visDocPreviewScope);
        public static Src Layers_Snap => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerSnap);
        public static Src Layers_ColorTrans => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColorTrans);
        public static Src Layers_Visible => new Src(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerVisible);

        //text transform
        public static Src TxtAngle => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormAngle);
        public static Src TxtHeight => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormHeight);
        public static Src TxtLocPinX => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinX);
        public static Src TxtLocPinY => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinY);
        public static Src TxtPinX => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinX);
        public static Src TxtPinY => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinY);
        public static Src TxtWidth => new Src(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormWidth);

        // user defined cells
        public static Src User_Prompt => new Src(SEC.visSectionUser, ROW.visRowUser, CEL.visUserPrompt);
        public static Src User_Value => new Src(SEC.visSectionUser, ROW.visRowUser, CEL.visUserValue);

        // Fields
        public static Src Fields_Calendar => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCalendar);
        public static Src Fields_Format => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldFormat);
        public static Src Fields_ObjectKind => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldObjectKind);
        public static Src Fields_Type => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldType);
        public static Src Fields_UICat => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICategory);
        public static Src Fields_UICod => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICode);
        public static Src Fields_UIFmt => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUIFormat);
        public static Src Fields_Value => new Src(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCell);

        // text block format
        public static Src BottomMargin => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBottomMargin);
        public static Src DefaultTabStop => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDefaultTabStop);
        public static Src LeftMargin => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkLeftMargin);
        public static Src RightMargin => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkRightMargin);
        public static Src TextBkgnd => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgnd);
        public static Src TextBkgndTrans => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgndTrans);
        public static Src TextDirection => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDirection);
        public static Src TopMargin => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkTopMargin);
        public static Src VerticalAlign => new Src(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkVerticalAlign);

        // Action tags
        public static Src SmartTags_ButtonFace => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagButtonFace);
        public static Src SmartTags_Description => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDescription);
        public static Src SmartTags_Disabled => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisabled);
        public static Src SmartTags_DisplayMode => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisplayMode);
        public static Src SmartTags_TagName => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagName);
        public static Src SmartTags_X => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagX);
        public static Src SmartTags_XJustify => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagXJustify);
        public static Src SmartTags_Y => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagY);
        public static Src SmartTags_YJustify => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagYJustify);

        // style
        public static Src EnableFillProps => new Src(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesFill);
        public static Src EnableLineProps => new Src(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesLine);
        public static Src EnableTextProps => new Src(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesText);

        //tabs
        public static Src Tabs_Alignment => new Src(SEC.visSectionTab, ROW.visRowTab, CEL.visTabAlign);
        public static Src Tabs_Position => new Src(SEC.visSectionTab, ROW.visRowTab, CEL.visTabPos);
        public static Src Tabs_StopCount => new Src(SEC.visSectionTab, ROW.visRowTab, CEL.visTabStopCount);

        // shape layout
        public static Src ConFixedCode => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOConFixedCode);
        public static Src ConLineJumpCode => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpCode);
        public static Src ConLineJumpDirX => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirX);
        public static Src ConLineJumpDirY => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirY);
        public static Src ConLineJumpStyle => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpStyle);
        public static Src ConLineRouteExt => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOLineRouteExt);
        public static Src ShapeFixedCode => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOFixedCode);
        public static Src ShapePermeablePlace => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermeablePlace);
        public static Src ShapePermeableX => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermX);
        public static Src ShapePermeableY => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermY);
        public static Src ShapePlaceFlip => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceFlip);
        public static Src ShapePlaceStyle => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceStyle);
        public static Src ShapePlowCode => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlowCode);
        public static Src ShapeRouteStyle => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORouteStyle);
        public static Src ShapeSplit => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplit);
        public static Src ShapeSplittable => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplittable);
        public static Src DisplayLevel => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLODisplayLevel); // new in Visio 2010
        public static Src Relationships => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORelationships); // new in Visio 2010
    }
}
