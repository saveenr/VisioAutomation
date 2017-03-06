using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;

namespace VisioAutomation.ShapeSheet
{
    public static class SrcConstants
    {
        // Actions
        private static Src ActionCell(CEL c) => new Src(SEC.visSectionAction, ROW.visRowAction, c);
        public static Src ActionAction => ActionCell(CEL.visActionAction);
        public static Src ActionBeginGroup => ActionCell(CEL.visActionBeginGroup);
        public static Src ActionButtonFace => ActionCell(CEL.visActionButtonFace);
        public static Src ActionChecked => ActionCell(CEL.visActionChecked);
        public static Src ActionDisabled => ActionCell(CEL.visActionDisabled);
        public static Src ActionInvisible => ActionCell(CEL.visActionInvisible);
        public static Src ActionMenu => ActionCell(CEL.visActionMenu);
        public static Src ActionReadOnly => ActionCell(CEL.visActionReadOnly);
        public static Src ActionSortKey => ActionCell(CEL.visActionSortKey);
        public static Src ActionTagName => ActionCell(CEL.visActionTagName);
        public static Src ActionFlyoutChild => ActionCell(CEL.visActionFlyoutChild); // new for visio 2010

        // Alignment
        private static Src AlignCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowAlign, c);
        public static Src AlignBottom => AlignCell(CEL.visAlignBottom);
        public static Src AlignCenter => AlignCell(CEL.visAlignCenter);
        public static Src AlignLeft => AlignCell(CEL.visAlignLeft);
        public static Src AlignMiddle => AlignCell(CEL.visAlignMiddle);
        public static Src AlignRight => AlignCell(CEL.visAlignRight);
        public static Src AlignTop => AlignCell(CEL.visAlignTop);

        // Annotation
        private static Src AnnotationCell(CEL c) => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, c);
        public static Src AnnotationComment => AnnotationCell(CEL.visAnnotationComment);
        public static Src AnnotationDate => AnnotationCell(CEL.visAnnotationDate);
        public static Src AnnotationLangID => AnnotationCell(CEL.visAnnotationLangID);
        public static Src AnnotationMarkerIndex => AnnotationCell(CEL.visAnnotationMarkerIndex);
        public static Src AnnotationX => AnnotationCell(CEL.visAnnotationX);
        public static Src AnnotationY => AnnotationCell(CEL.visAnnotationY);

        // Character
        private static Src CharCell(CEL c) => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, c);
        public static Src CharAsianFont => CharCell(CEL.visCharacterAsianFont);
        public static Src CharCase => CharCell(CEL.visCharacterCase);
        public static Src CharColor => CharCell(CEL.visCharacterColor);
        public static Src CharComplexScriptFont => CharCell(CEL.visCharacterComplexScriptFont);
        public static Src CharComplexScriptSize => CharCell(CEL.visCharacterComplexScriptSize);
        public static Src CharDoubleStrikethrough => CharCell(CEL.visCharacterDoubleStrikethrough);
        public static Src CharDoubleUnderline => CharCell(CEL.visCharacterDblUnderline);
        public static Src CharFont => CharCell(CEL.visCharacterFont);
        public static Src CharLangID => CharCell(CEL.visCharacterLangID);
        public static Src CharLocale => CharCell(CEL.visCharacterLocale);
        public static Src CharLocalizeFont => CharCell(CEL.visCharacterLocalizeFont);
        public static Src CharOverline => CharCell(CEL.visCharacterOverline);
        public static Src CharPerpendicular => CharCell(CEL.visCharacterPerpendicular);
        public static Src CharPos => CharCell(CEL.visCharacterPos);
        public static Src CharRTLText => CharCell(CEL.visCharacterRTLText);
        public static Src CharFontScale => CharCell(CEL.visCharacterFontScale);
        public static Src CharSize => CharCell(CEL.visCharacterSize);
        public static Src CharLetterspace => CharCell(CEL.visCharacterLetterspace);
        public static Src CharStrikethru => CharCell(CEL.visCharacterStrikethru);
        public static Src CharStyle => CharCell(CEL.visCharacterStyle);
        public static Src CharColorTransparency => CharCell(CEL.visCharacterColorTrans);
        public static Src CharUseVertical => CharCell(CEL.visCharacterUseVertical);

        // Connections
        private static Src ConnectionCell(CEL c) => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, c);
        public static Src ConnectionD => ConnectionCell(CEL.visCnnctD);
        public static Src ConnectionDirX => ConnectionCell(CEL.visCnnctDirX);
        public static Src ConnectionDirY => ConnectionCell(CEL.visCnnctDirY);
        public static Src ConnectionType => ConnectionCell(CEL.visCnnctType);
        public static Src ConnectionX => ConnectionCell(CEL.visX);
        public static Src ConnectionY => ConnectionCell(CEL.visY);

        // Controls
        private static Src ControlCell(CEL c) => new Src(SEC.visSectionControls, ROW.visRowControl, c);
        public static Src ControlCanGlue => ControlCell(CEL.visCtlGlue);
        public static Src ControlTip => ControlCell(CEL.visCtlTip);
        public static Src ControlXCon => ControlCell(CEL.visCtlXCon);
        public static Src ControlX => ControlCell(CEL.visCtlX);
        public static Src ControlXDyn => ControlCell(CEL.visCtlXDyn);
        public static Src ControlYCon => ControlCell(CEL.visCtlYCon);
        public static Src ControlY => ControlCell(CEL.visCtlY);
        public static Src ControlYDyn => ControlCell(CEL.visCtlYDyn);

        // Document Properties
        private static Src DocCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowDoc, c);
        public static Src DocAddMarkup => DocCell(CEL.visDocAddMarkup);
        public static Src DocLangID => DocCell(CEL.visDocLangID);
        public static Src DocLockPreview => DocCell(CEL.visDocLockPreview);
        public static Src DocOutputFormat => DocCell(CEL.visDocOutputFormat);
        public static Src DocPreviewQuality => DocCell(CEL.visDocPreviewQuality);
        public static Src DocPreviewScope => DocCell(CEL.visDocPreviewScope);
        public static Src DocViewMarkup => DocCell(CEL.visDocViewMarkup);

        // Events
        private static Src EventCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowEvent, c);
        public static Src EventDblClick => EventCell(CEL.visEvtCellDblClick);
        public static Src EventDrop => EventCell(CEL.visEvtCellDrop);
        public static Src EventMultiDrop => EventCell(CEL.visEvtCellMultiDrop);
        public static Src EventXFMod => EventCell(CEL.visEvtCellXFMod);
        public static Src EventTheText => EventCell(CEL.visEvtCellTheText);

        // ForeignImageInfo
        private static Src ForeignImgCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowForeign, c);
        public static Src ForeignImageHeight => ForeignImgCell(CEL.visFrgnImgHeight);
        public static Src ForeignImageOffsetX => ForeignImgCell(CEL.visFrgnImgOffsetX);
        public static Src ForeignImageOffsetY => ForeignImgCell(CEL.visFrgnImgOffsetY);
        public static Src ForeignImageWidth => ForeignImgCell(CEL.visFrgnImgWidth);

        // Geometry 
        private static Src GeometryVertexCell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, c);
        public static Src GeometryVertexA => GeometryVertexCell(CEL.visBow);
        public static Src GeometryVertexB => GeometryVertexCell(CEL.visControlX);
        public static Src GeometryVertexC => GeometryVertexCell(CEL.visEccentricityAngle);
        public static Src GeometryVertexD => GeometryVertexCell(CEL.visAspectRatio);
        public static Src GeometryVertexE => GeometryVertexCell(CEL.visNURBSData);
        public static Src GeometryVertexX => GeometryVertexCell(CEL.visX);
        public static Src GeometryVertexY => GeometryVertexCell(CEL.visY);

        // Geometry
        private static Src GeometryRowCell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, c);
        public static Src GeometryNoFill => GeometryRowCell(CEL.visCompNoFill);
        public static Src GeometryNoLine => GeometryRowCell(CEL.visCompNoLine);
        public static Src GeometryNoShow => GeometryRowCell(CEL.visCompNoShow);
        public static Src GeometryNoSnap => GeometryRowCell(CEL.visCompNoSnap);
        public static Src GeometryNoQuickDrag => GeometryRowCell(CEL.visCompNoQuickDrag);

        // Fill Format
        private static Src FillCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowFill, c);
        public static Src FillBackground => FillCell(CEL.visFillBkgnd);
        public static Src FillBackgroundTransparency => FillCell(CEL.visFillBkgndTrans);
        public static Src FillForeground => FillCell(CEL.visFillForegnd);
        public static Src FillForegroundTransparency => FillCell(CEL.visFillForegndTrans);
        public static Src FillPattern => FillCell(CEL.visFillPattern);
        public static Src FillShadowObliqueAngle => FillCell(CEL.visFillShdwObliqueAngle);
        public static Src FillShadowOffsetX => FillCell(CEL.visFillShdwOffsetX);
        public static Src FillShadowOffsetY => FillCell(CEL.visFillShdwOffsetY);
        public static Src FillShadowScaleFactor => FillCell(CEL.visFillShdwScaleFactor);
        public static Src FillShadowType => FillCell(CEL.visFillShdwType);
        public static Src FillShadowBackground => FillCell(CEL.visFillShdwBkgnd);
        public static Src FillShadowBackgroundTransparency => FillCell(CEL.visFillShdwBkgndTrans);
        public static Src FillShadowForeground => FillCell(CEL.visFillShdwForegnd);
        public static Src FillShadowForegroundTransparency => FillCell(CEL.visFillShdwForegndTrans);
        public static Src FillShadowPattern => FillCell(CEL.visFillShdwPattern);

        // GlueInfo
        private static Src GlueCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src GlueBeginTrigger => GlueCell(CEL.visBegTrigger);
        public static Src GlueEndTrigger => GlueCell(CEL.visEndTrigger);
        public static Src GlueType => GlueCell(CEL.visGlueType);
        public static Src GlueWalkPref => GlueCell(CEL.visWalkPref);

        // GroupProperties
        private static Src GroupCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowGroup, c);
        public static Src GroupDisplayMode => GroupCell(CEL.visGroupDisplayMode);
        public static Src GroupDontMoveChildren => GroupCell(CEL.visGroupDontMoveChildren);
        public static Src GroupIsDropTarget => GroupCell(CEL.visGroupIsDropTarget);
        public static Src GroupIsSnapTarget => GroupCell(CEL.visGroupIsSnapTarget);
        public static Src GroupIsTextEditTarget => GroupCell(CEL.visGroupIsTextEditTarget);
        public static Src GroupSelectMode => GroupCell(CEL.visGroupSelectMode);

        // Hyperlinks
        private static Src HyperlinkCell(CEL c) => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, c);
        public static Src HyperlinkAddress => HyperlinkCell(CEL.visHLinkAddress);
        public static Src HyperlinkDefault => HyperlinkCell(CEL.visHLinkDefault);
        public static Src HyperlinkDescription => HyperlinkCell(CEL.visHLinkDescription);
        public static Src HyperlinkExtraInfo => HyperlinkCell(CEL.visHLinkExtraInfo);
        public static Src HyperlinkFrame => HyperlinkCell(CEL.visHLinkFrame);
        public static Src HyperlinkInvisible => HyperlinkCell(CEL.visHLinkInvisible);
        public static Src HyperlinkNewWindow => HyperlinkCell(CEL.visHLinkNewWin);
        public static Src HyperlinkSortKey => HyperlinkCell(CEL.visHLinkSortKey);
        public static Src HyperlinkSubAddress => HyperlinkCell(CEL.visHLinkSubAddress);

        // Image Properties
        private static Src ImageCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowImage, c);
        public static Src ImageBlur => ImageCell(CEL.visImageBlur);
        public static Src ImageBrightness => ImageCell(CEL.visImageBrightness);
        public static Src ImageContrast => ImageCell(CEL.visImageContrast);
        public static Src ImageDenoise => ImageCell(CEL.visImageDenoise);
        public static Src ImageGamma => ImageCell(CEL.visImageGamma);
        public static Src ImageSharpen => ImageCell(CEL.visImageSharpen);
        public static Src ImageTransparency => ImageCell(CEL.visImageTransparency);

        // Line format
        private static Src LineCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowLine, c);
        public static Src LineBeginArrow => LineCell(CEL.visLineBeginArrow);
        public static Src LineBeginArrowSize => LineCell(CEL.visLineBeginArrowSize);
        public static Src LineEndArrow => LineCell(CEL.visLineEndArrow);
        public static Src LineEndArrowSize => LineCell(CEL.visLineEndArrowSize);
        public static Src LineCap => LineCell(CEL.visLineEndCap);
        public static Src LineColor => LineCell(CEL.visLineColor);
        public static Src LineColorTransparency => LineCell(CEL.visLineColorTrans);
        public static Src LinePattern => LineCell(CEL.visLinePattern);
        public static Src LineWeight => LineCell(CEL.visLineWeight);
        public static Src LineRounding => LineCell(CEL.visLineRounding);

        // Miscellaneous
        private static Src MiscCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src MiscCalendar => MiscCell(CEL.visObjCalendar);
        public static Src MiscComment => MiscCell(CEL.visComment);
        public static Src MiscDropOnPageScale => MiscCell(CEL.visObjDropOnPageScale);
        public static Src MiscDynFeedback => MiscCell(CEL.visDynFeedback);
        public static Src MiscIsDropSource => MiscCell(CEL.visDropSource);
        public static Src MiscLangID => MiscCell(CEL.visObjLangID);
        public static Src MiscLocalizeMerge => MiscCell(CEL.visObjLocalizeMerge);
        public static Src MiscNoAlignBox => MiscCell(CEL.visNoAlignBox);
        public static Src MiscNoCtlHandles => MiscCell(CEL.visNoCtlHandles);
        public static Src MiscNoLiveDynamics => MiscCell(CEL.visNoLiveDynamics);
        public static Src MiscNonPrinting => MiscCell(CEL.visNonPrinting);
        public static Src MiscNoObjHandles => MiscCell(CEL.visNoObjHandles);
        public static Src MiscObjType => MiscCell(CEL.visLOFlags);
        public static Src MiscUpdateAlignBox => MiscCell(CEL.visUpdateAlignBox);
        public static Src MiscHideText => MiscCell(CEL.visHideText);

        // 1d endpoints
        private static Src OneDCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXForm1D, c);
        public static Src OneDBeginX => OneDCell(CEL.vis1DBeginX);
        public static Src OneDBeginY => OneDCell(CEL.vis1DBeginY);
        public static Src OneDEndX => OneDCell(CEL.vis1DEndX);
        public static Src OneDEndY => OneDCell(CEL.vis1DEndY);

        // page layout
        private static Src PageLayoutCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPageLayout, c);
        public static Src PageLayoutAvenueSizeX => PageLayoutCell(CEL.visPLOAvenueSizeX);
        public static Src PageLayoutAvenueSizeY => PageLayoutCell(CEL.visPLOAvenueSizeY);
        public static Src PageLayoutBlockSizeX => PageLayoutCell(CEL.visPLOBlockSizeX);
        public static Src PageLayoutBlockSizeY => PageLayoutCell(CEL.visPLOBlockSizeY);
        public static Src PageLayoutCtrlAsInput => PageLayoutCell(CEL.visPLOCtrlAsInput);
        public static Src PageLayoutDynamicsOff => PageLayoutCell(CEL.visPLODynamicsOff);
        public static Src PageLayoutEnableGrid => PageLayoutCell(CEL.visPLOEnableGrid);
        public static Src PageLayoutLineAdjustFrom => PageLayoutCell(CEL.visPLOLineAdjustFrom);
        public static Src PageLayoutLineAdjustTo => PageLayoutCell(CEL.visPLOLineAdjustTo);
        public static Src PageLayoutLineJumpCode => PageLayoutCell(CEL.visPLOJumpCode);
        public static Src PageLayoutLineJumpFactorX => PageLayoutCell(CEL.visPLOJumpFactorX);
        public static Src PageLayoutLineJumpFactorY => PageLayoutCell(CEL.visPLOJumpFactorY);
        public static Src PageLayoutLineJumpStyle => PageLayoutCell(CEL.visPLOJumpStyle);
        public static Src PageLayoutLineRouteExt => PageLayoutCell(CEL.visPLOLineRouteExt);
        public static Src PageLayoutLineToLineX => PageLayoutCell(CEL.visPLOLineToLineX);
        public static Src PageLayoutLineToLineY => PageLayoutCell(CEL.visPLOLineToLineY);
        public static Src PageLayoutLineToNodeX => PageLayoutCell(CEL.visPLOLineToNodeX);
        public static Src PageLayoutLineToNodeY => PageLayoutCell(CEL.visPLOLineToNodeY);
        public static Src PageLayoutLineJumpDirX => PageLayoutCell(CEL.visPLOJumpDirX);
        public static Src PageLayoutLineJumpDirY => PageLayoutCell(CEL.visPLOJumpDirY);
        public static Src PageLayoutPageShapeSplit => PageLayoutCell(CEL.visPLOSplit);
        public static Src PageLayoutPlaceDepth => PageLayoutCell(CEL.visPLOPlaceDepth);
        public static Src PageLayoutPlaceFlip => PageLayoutCell(CEL.visPLOPlaceFlip);
        public static Src PageLayoutPlaceStyle => PageLayoutCell(CEL.visPLOPlaceStyle);
        public static Src PageLayoutPlowCode => PageLayoutCell(CEL.visPLOPlowCode);
        public static Src PageLayoutResizePage => PageLayoutCell(CEL.visPLOResizePage);
        public static Src PageLayoutRouteStyle => PageLayoutCell(CEL.visPLORouteStyle);
        public static Src PageLayoutAvoidPageBreaks => PageLayoutCell(CEL.visPLOAvoidPageBreaks); // new in Visio 2010

        // print properties
        private static Src PrintCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, c);
        public static Src PrintLeftMargin => PrintCell(CEL.visPrintPropertiesLeftMargin);
        public static Src PrintCenterX => PrintCell(CEL.visPrintPropertiesCenterX);
        public static Src PrintCenterY => PrintCell(CEL.visPrintPropertiesCenterY);
        public static Src PrintOnPage => PrintCell(CEL.visPrintPropertiesOnPage);
        public static Src PrintBottomMargin => PrintCell(CEL.visPrintPropertiesBottomMargin);
        public static Src PrintRightMargin => PrintCell(CEL.visPrintPropertiesRightMargin);
        public static Src PrintPagesX => PrintCell(CEL.visPrintPropertiesPagesX);
        public static Src PrintPagesY => PrintCell(CEL.visPrintPropertiesPagesY);
        public static Src PrintTopMargin => PrintCell(CEL.visPrintPropertiesTopMargin);
        public static Src PrintPaperKind => PrintCell(CEL.visPrintPropertiesPaperKind);
        public static Src PrintGrid => PrintCell(CEL.visPrintPropertiesPrintGrid);
        public static Src PrintPageOrientation => PrintCell(CEL.visPrintPropertiesPageOrientation);
        public static Src PrintScaleX => PrintCell(CEL.visPrintPropertiesScaleX);
        public static Src PrintScaleY => PrintCell(CEL.visPrintPropertiesScaleY);
        public static Src PrintPaperSource => PrintCell(CEL.visPrintPropertiesPaperSource);

        // page properties
        private static Src PagePropCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPage, c);
        public static Src PageDrawingScale => PagePropCell(CEL.visPageDrawingScale);
        public static Src PageDrawingScaleType => PagePropCell(CEL.visPageDrawScaleType);
        public static Src PageDrawingSizeType => PagePropCell(CEL.visPageDrawSizeType);
        public static Src PageInhibitSnap => PagePropCell(CEL.visPageInhibitSnap);
        public static Src PageHeight => PagePropCell(CEL.visPageHeight);
        public static Src PageScale => PagePropCell(CEL.visPageScale);
        public static Src PageWidth => PagePropCell(CEL.visPageWidth);
        public static Src PageShadowObliqueAngle => PagePropCell(CEL.visPageShdwObliqueAngle);
        public static Src PageShadowOffsetX => PagePropCell(CEL.visPageShdwOffsetX);
        public static Src PageShadowOffsetY => PagePropCell(CEL.visPageShdwOffsetY);
        public static Src PageShadowScaleFactor => PagePropCell(CEL.visPageShdwScaleFactor);
        public static Src PageShadowType => PagePropCell(CEL.visPageShdwType);
        public static Src PageUIVisibility => PagePropCell(CEL.visPageUIVisibility);
        public static Src PageDrawingResizeType => PagePropCell(CEL.visPageDrawResizeType); // new in Visio 2010

        // paragraph
        private static Src ParaCell(CEL c) => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, c);
        public static Src ParaBullet => ParaCell(CEL.visBulletIndex);
        public static Src ParaBulletFont => ParaCell(CEL.visBulletFont);
        public static Src ParaBulletFontSize => ParaCell(CEL.visBulletFontSize);
        public static Src ParaBulletStr => ParaCell(CEL.visBulletString);
        public static Src ParaFlags => ParaCell(CEL.visFlags);
        public static Src ParaHorizontalAlign => ParaCell(CEL.visHorzAlign);
        public static Src ParaIndentFirst => ParaCell(CEL.visIndentFirst);
        public static Src ParaIndentLeft => ParaCell(CEL.visIndentLeft);
        public static Src ParaIndentRight => ParaCell(CEL.visIndentRight);
        public static Src ParaLocalizeBulletFont => ParaCell(CEL.visLocalizeBulletFont);
        public static Src ParaSpacingAfter => ParaCell(CEL.visSpaceAfter);
        public static Src ParaSpacingBefore => ParaCell(CEL.visSpaceBefore);
        public static Src ParaSpacingLine => ParaCell(CEL.visSpaceLine);
        public static Src ParaTextPosAfterBullet => ParaCell(CEL.visTextPosAfterBullet);

        // protection
        private static Src LockCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowLock, c);
        public static Src LockAspect => LockCell(CEL.visLockAspect);
        public static Src LockBegin => LockCell(CEL.visLockBegin);
        public static Src LockCalcWH => LockCell(CEL.visLockCalcWH);
        public static Src LockCrop => LockCell(CEL.visLockCrop);
        public static Src LockCustProp => LockCell(CEL.visLockCustProp);
        public static Src LockDelete => LockCell(CEL.visLockDelete);
        public static Src LockEnd => LockCell(CEL.visLockEnd);
        public static Src LockFormat => LockCell(CEL.visLockFormat);
        public static Src LockFromGroupFormat => LockCell(CEL.visLockFromGroupFormat);
        public static Src LockGroup => LockCell(CEL.visLockGroup);
        public static Src LockHeight => LockCell(CEL.visLockHeight);
        public static Src LockMoveX => LockCell(CEL.visLockMoveX);
        public static Src LockMoveY => LockCell(CEL.visLockMoveY);
        public static Src LockRotate => LockCell(CEL.visLockRotate);
        public static Src LockSelect => LockCell(CEL.visLockSelect);
        public static Src LockTextEdit => LockCell(CEL.visLockTextEdit);
        public static Src LockThemeColors => LockCell(CEL.visLockThemeColors);
        public static Src LockThemeEffects => LockCell(CEL.visLockThemeEffects);
        public static Src LockVtxEdit => LockCell(CEL.visLockVtxEdit);
        public static Src LockWidth => LockCell(CEL.visLockWidth);

        // ruler and grid
        private static Src RulerGridCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, c);
        public static Src XGridDensity => RulerGridCell(CEL.visXGridDensity);
        public static Src XGridOrigin => RulerGridCell(CEL.visXGridOrigin);
        public static Src XGridSpacing => RulerGridCell(CEL.visXGridSpacing);
        public static Src YGridDensity => RulerGridCell(CEL.visYGridDensity);
        public static Src YGridOrigin => RulerGridCell(CEL.visYGridOrigin);
        public static Src YGridSpacing => RulerGridCell(CEL.visYGridSpacing);
        public static Src XRulerDensity => RulerGridCell(CEL.visXRulerDensity);
        public static Src XRulerOrigin => RulerGridCell(CEL.visXRulerOrigin);
        public static Src YRulerDensity => RulerGridCell(CEL.visYRulerDensity);
        public static Src YRulerOrigin => RulerGridCell(CEL.visYRulerOrigin);

        // Shape Tranform
        private static Src XFormCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXFormOut, c);
        public static Src XFormAngle => XFormCell(CEL.visXFormAngle);
        public static Src XFormFlipX => XFormCell(CEL.visXFormFlipX);
        public static Src XFormFlipY => XFormCell(CEL.visXFormFlipY);
        public static Src XFormHeight => XFormCell(CEL.visXFormHeight);
        public static Src XFormLocPinX => XFormCell(CEL.visXFormLocPinX);
        public static Src XFormLocPinY => XFormCell(CEL.visXFormLocPinY);
        public static Src XFormPinX => XFormCell(CEL.visXFormPinX);
        public static Src XFormPinY => XFormCell(CEL.visXFormPinY);
        public static Src XFormResizeMode => XFormCell(CEL.visXFormResizeMode);
        public static Src XFormWidth => XFormCell(CEL.visXFormWidth);

        // reviewer
        private static Src ReviewerCell(CEL c) => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, c);
        public static Src ReviewerColor => ReviewerCell(CEL.visReviewerColor);
        public static Src ReviewerInitials => ReviewerCell(CEL.visReviewerInitials);
        public static Src ReviewerName => ReviewerCell(CEL.visReviewerName);

        // shape data
        private static Src CustPropCell(CEL c) => new Src(SEC.visSectionProp, ROW.visRowProp, c);
        public static Src CustPropSortKey => CustPropCell(CEL.visCustPropsSortKey);
        public static Src CustPropAsk => CustPropCell(CEL.visCustPropsAsk);
        public static Src CustPropCalendar => CustPropCell(CEL.visCustPropsCalendar);
        public static Src CustPropFormat => CustPropCell(CEL.visCustPropsFormat);
        public static Src CustPropInvisible => CustPropCell(CEL.visCustPropsInvis);
        public static Src CustPropLabel => CustPropCell(CEL.visCustPropsLabel);
        public static Src CustPropLangId => CustPropCell(CEL.visCustPropsLangID);
        public static Src CustPropPrompt => CustPropCell(CEL.visCustPropsPrompt);
        public static Src CustPropType => CustPropCell(CEL.visCustPropsType);
        public static Src CustPropValue => CustPropCell(CEL.visCustPropsValue);

        // Layers
        private static Src LayerCell(CEL c) => new Src(SEC.visSectionLayer, ROW.visRowLayer, c);
        public static Src LayerActive => LayerCell(CEL.visLayerActive);
        public static Src LayerColor => LayerCell(CEL.visLayerColor);
        public static Src LayerGlue => LayerCell(CEL.visLayerGlue);
        public static Src LayerLocked => LayerCell(CEL.visLayerLock);
        public static Src LayerPrint => LayerCell(CEL.visDocPreviewScope);
        public static Src LayerSnap => LayerCell(CEL.visLayerSnap);
        public static Src LayerColorTransparency => LayerCell(CEL.visLayerColorTrans);
        public static Src LayerVisible => LayerCell(CEL.visLayerVisible);

        //text transform
        private static Src TextXFormCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowTextXForm, c);
        public static Src TextXFormAngle => TextXFormCell(CEL.visXFormAngle);
        public static Src TextXFormHeight => TextXFormCell(CEL.visXFormHeight);
        public static Src TextXFormLocPinX => TextXFormCell(CEL.visXFormLocPinX);
        public static Src TextXFormLocPinY => TextXFormCell(CEL.visXFormLocPinY);
        public static Src TextXFormPinX => TextXFormCell(CEL.visXFormPinX);
        public static Src TextXFormPinY => TextXFormCell(CEL.visXFormPinY);
        public static Src TextXFormWidth => TextXFormCell(CEL.visXFormWidth);

        // user defined cells
        private static Src UserDefCell(CEL c) => new Src(SEC.visSectionUser, ROW.visRowUser, c);
        public static Src UserDefCellPrompt => UserDefCell(CEL.visUserPrompt);
        public static Src UserDelCellValue => UserDefCell(CEL.visUserValue);

        // Fields
        private static Src FieldCell(CEL c) => new Src(SEC.visSectionTextField, ROW.visRowField, c);
        public static Src FieldCalendar => FieldCell(CEL.visFieldCalendar);
        public static Src FieldFormat => FieldCell(CEL.visFieldFormat);
        public static Src FieldObjectKind => FieldCell(CEL.visFieldObjectKind);
        public static Src FieldType => FieldCell(CEL.visFieldType);
        public static Src FieldUICategory => FieldCell(CEL.visFieldUICategory);
        public static Src FieldUICode => FieldCell(CEL.visFieldUICode);
        public static Src FieldUIFormat => FieldCell(CEL.visFieldUIFormat);
        public static Src FieldValue => FieldCell(CEL.visFieldCell);

        // text block format
        private static Src TextBlockCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowText, c);
        public static Src TextBlockBottomMargin => TextBlockCell(CEL.visTxtBlkBottomMargin);
        public static Src TextBlockDefaultTabStop => TextBlockCell(CEL.visTxtBlkDefaultTabStop);
        public static Src TextBlockLeftMargin => TextBlockCell(CEL.visTxtBlkLeftMargin);
        public static Src TextBlockRightMargin => TextBlockCell(CEL.visTxtBlkRightMargin);
        public static Src TextBlockBackground => TextBlockCell(CEL.visTxtBlkBkgnd);
        public static Src TextBlockBackgroundTransparency => TextBlockCell(CEL.visTxtBlkBkgndTrans);
        public static Src TextBlockDirection => TextBlockCell(CEL.visTxtBlkDirection);
        public static Src TextBlockTopMargin => TextBlockCell(CEL.visTxtBlkTopMargin);
        public static Src TextBlockVerticalAlign => TextBlockCell(CEL.visTxtBlkVerticalAlign);

        // Action tags
        private static Src SmartTagCell(CEL c) => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, c);
        public static Src SmartTagButtonFace => SmartTagCell(CEL.visSmartTagButtonFace);
        public static Src SmartTagDescription => SmartTagCell(CEL.visSmartTagDescription);
        public static Src SmartTagDisabled => SmartTagCell(CEL.visSmartTagDisabled);
        public static Src SmartTagDisplayMode => SmartTagCell(CEL.visSmartTagDisplayMode);
        public static Src SmartTagTagName => SmartTagCell(CEL.visSmartTagName);
        public static Src SmartTagX => SmartTagCell(CEL.visSmartTagX);
        public static Src SmartTagXJustify => SmartTagCell(CEL.visSmartTagXJustify);
        public static Src SmartTagY => SmartTagCell(CEL.visSmartTagY);
        public static Src SmartTagYJustify => SmartTagCell(CEL.visSmartTagYJustify);

        // style
        private static Src StyleCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowStyle, c);
        public static Src StyleEnableFillProps => StyleCell(CEL.visStyleIncludesFill);
        public static Src StyleEnableLineProps => StyleCell(CEL.visStyleIncludesLine);
        public static Src StyleEnableTextProps => StyleCell(CEL.visStyleIncludesText);

        //tabs
        private static Src TabCell(CEL c) => new Src(SEC.visSectionTab, ROW.visRowTab, c);
        public static Src TabAlignment => TabCell(CEL.visTabAlign);
        public static Src TabPosition => TabCell(CEL.visTabPos);
        public static Src TabStopCount => TabCell(CEL.visTabStopCount);

        // shape layout
        private static Src ShapeLayoutCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, c);
        public static Src ShapeLayoutConFixedCode => ShapeLayoutCell(CEL.visSLOConFixedCode);
        public static Src ShapeLayoutConLineJumpCode => ShapeLayoutCell(CEL.visSLOJumpCode);
        public static Src ShapeLayoutConLineJumpDirX => ShapeLayoutCell(CEL.visSLOJumpDirX);
        public static Src ShapeLayoutConLineJumpDirY => ShapeLayoutCell(CEL.visSLOJumpDirY);
        public static Src ShapeLayoutConLineJumpStyle => ShapeLayoutCell(CEL.visSLOJumpStyle);
        public static Src ShapeLayoutConLineRouteExt => ShapeLayoutCell(CEL.visSLOLineRouteExt);
        public static Src ShapeLayoutFixedCode => ShapeLayoutCell(CEL.visSLOFixedCode);
        public static Src ShapeLayoutPermeablePlace => ShapeLayoutCell(CEL.visSLOPermeablePlace);
        public static Src ShapeLayoutPermeableX => ShapeLayoutCell(CEL.visSLOPermX);
        public static Src ShapeLayoutPermeableY => ShapeLayoutCell(CEL.visSLOPermY);
        public static Src ShapeLayoutPlaceFlip => ShapeLayoutCell(CEL.visSLOPlaceFlip);
        public static Src ShapeLayoutPlaceStyle => ShapeLayoutCell(CEL.visSLOPlaceStyle);
        public static Src ShapeLayoutPlowCode => ShapeLayoutCell(CEL.visSLOPlowCode);
        public static Src ShapeLayoutRouteStyle => ShapeLayoutCell(CEL.visSLORouteStyle);
        public static Src ShapeLayoutSplit => ShapeLayoutCell(CEL.visSLOSplit);
        public static Src ShapeLayoutSplittable => ShapeLayoutCell(CEL.visSLOSplittable);
        public static Src ShapeLayoutDisplayLevel => ShapeLayoutCell(CEL.visSLODisplayLevel); // new in Visio 2010
        public static Src ShapeLayoutRelationships => ShapeLayoutCell(CEL.visSLORelationships); // new in Visio 2010
    }
}
