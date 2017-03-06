using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;

namespace VisioAutomation.ShapeSheet
{
    public static class SrcConstants
    {
        // Actions
        private static Src ActionCell(CEL c) => new Src(SEC.visSectionAction, ROW.visRowAction, c);
        public static Src Actions_Action => ActionCell(CEL.visActionAction);
        public static Src Actions_BeginGroup => ActionCell(CEL.visActionBeginGroup);
        public static Src Actions_ButtonFace => ActionCell(CEL.visActionButtonFace);
        public static Src Actions_Checked => ActionCell(CEL.visActionChecked);
        public static Src Actions_Disabled => ActionCell(CEL.visActionDisabled);
        public static Src Actions_Invisible => ActionCell(CEL.visActionInvisible);
        public static Src Actions_Menu => ActionCell(CEL.visActionMenu);
        public static Src Actions_ReadOnly => ActionCell(CEL.visActionReadOnly);
        public static Src Actions_SortKey => ActionCell(CEL.visActionSortKey);
        public static Src Actions_TagName => ActionCell(CEL.visActionTagName);
        public static Src Actions_FlyoutChild => ActionCell(CEL.visActionFlyoutChild); // new for visio 2010

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
        public static Src Annotation_Comment => AnnotationCell(CEL.visAnnotationComment);
        public static Src Annotation_Date => AnnotationCell(CEL.visAnnotationDate);
        public static Src Annotation_LangID => AnnotationCell(CEL.visAnnotationLangID);
        public static Src Annotation_MarkerIndex => AnnotationCell(CEL.visAnnotationMarkerIndex);
        public static Src Annotation_X => AnnotationCell(CEL.visAnnotationX);
        public static Src Annotation_Y => AnnotationCell(CEL.visAnnotationY);

        // Character
        private static Src CharCell(CEL c) => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, c);
        public static Src CharAsianFont => CharCell(CEL.visCharacterAsianFont);
        public static Src CharCase => CharCell(CEL.visCharacterCase);
        public static Src CharColor => CharCell(CEL.visCharacterColor);
        public static Src CharComplexScriptFont => CharCell(CEL.visCharacterComplexScriptFont);
        public static Src CharComplexScriptSize => CharCell(CEL.visCharacterComplexScriptSize);
        public static Src CharDoubleStrikethrough => CharCell(CEL.visCharacterDoubleStrikethrough);
        public static Src CharDblUnderline => CharCell(CEL.visCharacterDblUnderline);
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
        public static Src CharColorTrans => CharCell(CEL.visCharacterColorTrans);
        public static Src CharUseVertical => CharCell(CEL.visCharacterUseVertical);

        // Connections
        private static Src ConnectionsCell(CEL c) => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, c);
        public static Src Connections_D => ConnectionsCell(CEL.visCnnctD);
        public static Src Connections_DirX => ConnectionsCell(CEL.visCnnctDirX);
        public static Src Connections_DirY => ConnectionsCell(CEL.visCnnctDirY);
        public static Src Connections_Type => ConnectionsCell(CEL.visCnnctType);
        public static Src Connections_X => ConnectionsCell(CEL.visX);
        public static Src Connections_Y => ConnectionsCell(CEL.visY);

        // Controls
        private static Src ControlsCell(CEL c) => new Src(SEC.visSectionControls, ROW.visRowControl, c);
        public static Src Controls_CanGlue => ControlsCell(CEL.visCtlGlue);
        public static Src Controls_Tip => ControlsCell(CEL.visCtlTip);
        public static Src Controls_XCon => ControlsCell(CEL.visCtlXCon);
        public static Src Controls_X => ControlsCell(CEL.visCtlX);
        public static Src Controls_XDyn => ControlsCell(CEL.visCtlXDyn);
        public static Src Controls_YCon => ControlsCell(CEL.visCtlYCon);
        public static Src Controls_Y => ControlsCell(CEL.visCtlY);
        public static Src Controls_YDyn => ControlsCell(CEL.visCtlYDyn);

        // Document Properties
        private static Src DocCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowDoc, c);
        public static Src AddMarkup => DocCell(CEL.visDocAddMarkup);
        public static Src DocLangID => DocCell(CEL.visDocLangID);
        public static Src LockPreview => DocCell(CEL.visDocLockPreview);
        public static Src OutputFormat => DocCell(CEL.visDocOutputFormat);
        public static Src PreviewQuality => DocCell(CEL.visDocPreviewQuality);
        public static Src PreviewScope => DocCell(CEL.visDocPreviewScope);
        public static Src ViewMarkup => DocCell(CEL.visDocViewMarkup);

        // Events
        private static Src EventCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowEvent, c);
        public static Src EventDblClick => EventCell(CEL.visEvtCellDblClick);
        public static Src EventDrop => EventCell(CEL.visEvtCellDrop);
        public static Src EventMultiDrop => EventCell(CEL.visEvtCellMultiDrop);
        public static Src EventXFMod => EventCell(CEL.visEvtCellXFMod);
        public static Src TheText => EventCell(CEL.visEvtCellTheText);

        // ForeignImageInfo
        private static Src ImgCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowForeign, c);
        public static Src ImgHeight => ImgCell(CEL.visFrgnImgHeight);
        public static Src ImgOffsetX => ImgCell(CEL.visFrgnImgOffsetX);
        public static Src ImgOffsetY => ImgCell(CEL.visFrgnImgOffsetY);
        public static Src ImgWidth => ImgCell(CEL.visFrgnImgWidth);

        // Geometry 
        private static Src GeometryVertexCell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, c);
        public static Src Geometry_A => GeometryVertexCell(CEL.visBow);
        public static Src Geometry_B => GeometryVertexCell(CEL.visControlX);
        public static Src Geometry_C => GeometryVertexCell(CEL.visEccentricityAngle);
        public static Src Geometry_D => GeometryVertexCell(CEL.visAspectRatio);
        public static Src Geometry_E => GeometryVertexCell(CEL.visNURBSData);
        public static Src Geometry_X => GeometryVertexCell(CEL.visX);
        public static Src Geometry_Y => GeometryVertexCell(CEL.visY);

        // Geometry
        private static Src GeometryRowCell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, c);
        public static Src Geometry_NoFill => GeometryRowCell(CEL.visCompNoFill);
        public static Src Geometry_NoLine => GeometryRowCell(CEL.visCompNoLine);
        public static Src Geometry_NoShow => GeometryRowCell(CEL.visCompNoShow);
        public static Src Geometry_NoSnap => GeometryRowCell(CEL.visCompNoSnap);
        public static Src Geometry_NoQuickDrag => GeometryRowCell(CEL.visCompNoQuickDrag);

        // Fill Format
        private static Src FillCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowFill, c);
        public static Src FillBkgnd => FillCell(CEL.visFillBkgnd);
        public static Src FillBkgndTrans => FillCell(CEL.visFillBkgndTrans);
        public static Src FillForegnd => FillCell(CEL.visFillForegnd);
        public static Src FillForegndTrans => FillCell(CEL.visFillForegndTrans);
        public static Src FillPattern => FillCell(CEL.visFillPattern);
        public static Src ShapeShdwObliqueAngle => FillCell(CEL.visFillShdwObliqueAngle);
        public static Src ShapeShdwOffsetX => FillCell(CEL.visFillShdwOffsetX);
        public static Src ShapeShdwOffsetY => FillCell(CEL.visFillShdwOffsetY);
        public static Src ShapeShdwScaleFactor => FillCell(CEL.visFillShdwScaleFactor);
        public static Src ShapeShdwType => FillCell(CEL.visFillShdwType);
        public static Src ShdwBkgnd => FillCell(CEL.visFillShdwBkgnd);
        public static Src ShdwBkgndTrans => FillCell(CEL.visFillShdwBkgndTrans);
        public static Src ShdwForegnd => FillCell(CEL.visFillShdwForegnd);
        public static Src ShdwForegndTrans => FillCell(CEL.visFillShdwForegndTrans);
        public static Src ShdwPattern => FillCell(CEL.visFillShdwPattern);

        // GlueInfo
        private static Src GlueCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src BegTrigger => GlueCell(CEL.visBegTrigger);
        public static Src EndTrigger => GlueCell(CEL.visEndTrigger);
        public static Src GlueType => GlueCell(CEL.visGlueType);
        public static Src WalkPreference => GlueCell(CEL.visWalkPref);

        // GroupProperties
        private static Src GroupCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowGroup, c);
        public static Src DisplayMode => GroupCell(CEL.visGroupDisplayMode);
        public static Src DontMoveChildren => GroupCell(CEL.visGroupDontMoveChildren);
        public static Src IsDropTarget => GroupCell(CEL.visGroupIsDropTarget);
        public static Src IsSnapTarget => GroupCell(CEL.visGroupIsSnapTarget);
        public static Src IsTextEditTarget => GroupCell(CEL.visGroupIsTextEditTarget);
        public static Src SelectMode => GroupCell(CEL.visGroupSelectMode);

        // Hyperlinks
        private static Src HyperlinkCell(CEL c) => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, c);
        public static Src Hyperlink_Address => HyperlinkCell(CEL.visHLinkAddress);
        public static Src Hyperlink_Default => HyperlinkCell(CEL.visHLinkDefault);
        public static Src Hyperlink_Description => HyperlinkCell(CEL.visHLinkDescription);
        public static Src Hyperlink_ExtraInfo => HyperlinkCell(CEL.visHLinkExtraInfo);
        public static Src Hyperlink_Frame => HyperlinkCell(CEL.visHLinkFrame);
        public static Src Hyperlink_Invisible => HyperlinkCell(CEL.visHLinkInvisible);
        public static Src Hyperlink_NewWindow => HyperlinkCell(CEL.visHLinkNewWin);
        public static Src Hyperlink_SortKey => HyperlinkCell(CEL.visHLinkSortKey);
        public static Src Hyperlink_SubAddress => HyperlinkCell(CEL.visHLinkSubAddress);

        // Image Properties
        private static Src ImageCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowImage, c);
        public static Src Blur => ImageCell(CEL.visImageBlur);
        public static Src Brightness => ImageCell(CEL.visImageBrightness);
        public static Src Contrast => ImageCell(CEL.visImageContrast);
        public static Src Denoise => ImageCell(CEL.visImageDenoise);
        public static Src Gamma => ImageCell(CEL.visImageGamma);
        public static Src Sharpen => ImageCell(CEL.visImageSharpen);
        public static Src Transparency => ImageCell(CEL.visImageTransparency);

        // Line format
        private static Src ArrowCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowLine, c);
        public static Src BeginArrow => ArrowCell(CEL.visLineBeginArrow);
        public static Src BeginArrowSize => ArrowCell(CEL.visLineBeginArrowSize);
        public static Src EndArrow => ArrowCell(CEL.visLineEndArrow);
        public static Src EndArrowSize => ArrowCell(CEL.visLineEndArrowSize);
        public static Src LineCap => ArrowCell(CEL.visLineEndCap);
        public static Src LineColor => ArrowCell(CEL.visLineColor);
        public static Src LineColorTrans => ArrowCell(CEL.visLineColorTrans);
        public static Src LinePattern => ArrowCell(CEL.visLinePattern);
        public static Src LineWeight => ArrowCell(CEL.visLineWeight);
        public static Src Rounding => ArrowCell(CEL.visLineRounding);

        // Miscellaneous
        private static Src MiscCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src Calendar => MiscCell(CEL.visObjCalendar);
        public static Src Comment => MiscCell(CEL.visComment);
        public static Src DropOnPageScale => MiscCell(CEL.visObjDropOnPageScale);
        public static Src DynFeedback => MiscCell(CEL.visDynFeedback);
        public static Src IsDropSource => MiscCell(CEL.visDropSource);
        public static Src LangID => MiscCell(CEL.visObjLangID);
        public static Src LocalizeMerge => MiscCell(CEL.visObjLocalizeMerge);
        public static Src NoAlignBox => MiscCell(CEL.visNoAlignBox);
        public static Src NoCtlHandles => MiscCell(CEL.visNoCtlHandles);
        public static Src NoLiveDynamics => MiscCell(CEL.visNoLiveDynamics);
        public static Src NonPrinting => MiscCell(CEL.visNonPrinting);
        public static Src NoObjHandles => MiscCell(CEL.visNoObjHandles);
        public static Src ObjType => MiscCell(CEL.visLOFlags);
        public static Src UpdateAlignBox => MiscCell(CEL.visUpdateAlignBox);
        public static Src HideText => MiscCell(CEL.visHideText);

        // 1d endpoints
        private static Src OneDCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXForm1D, c);
        public static Src BeginX => OneDCell(CEL.vis1DBeginX);
        public static Src BeginY => OneDCell(CEL.vis1DBeginY);
        public static Src EndX => OneDCell(CEL.vis1DEndX);
        public static Src EndY => OneDCell(CEL.vis1DEndY);

        // page layout
        private static Src PageLayoutCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPageLayout, c);
        public static Src AvenueSizeX => PageLayoutCell(CEL.visPLOAvenueSizeX);
        public static Src AvenueSizeY => PageLayoutCell(CEL.visPLOAvenueSizeY);
        public static Src BlockSizeX => PageLayoutCell(CEL.visPLOBlockSizeX);
        public static Src BlockSizeY => PageLayoutCell(CEL.visPLOBlockSizeY);
        public static Src CtrlAsInput => PageLayoutCell(CEL.visPLOCtrlAsInput);
        public static Src DynamicsOff => PageLayoutCell(CEL.visPLODynamicsOff);
        public static Src EnableGrid => PageLayoutCell(CEL.visPLOEnableGrid);
        public static Src LineAdjustFrom => PageLayoutCell(CEL.visPLOLineAdjustFrom);
        public static Src LineAdjustTo => PageLayoutCell(CEL.visPLOLineAdjustTo);
        public static Src LineJumpCode => PageLayoutCell(CEL.visPLOJumpCode);
        public static Src LineJumpFactorX => PageLayoutCell(CEL.visPLOJumpFactorX);
        public static Src LineJumpFactorY => PageLayoutCell(CEL.visPLOJumpFactorY);
        public static Src LineJumpStyle => PageLayoutCell(CEL.visPLOJumpStyle);
        public static Src LineRouteExt => PageLayoutCell(CEL.visPLOLineRouteExt);
        public static Src LineToLineX => PageLayoutCell(CEL.visPLOLineToLineX);
        public static Src LineToLineY => PageLayoutCell(CEL.visPLOLineToLineY);
        public static Src LineToNodeX => PageLayoutCell(CEL.visPLOLineToNodeX);
        public static Src LineToNodeY => PageLayoutCell(CEL.visPLOLineToNodeY);
        public static Src PageLineJumpDirX => PageLayoutCell(CEL.visPLOJumpDirX);
        public static Src PageLineJumpDirY => PageLayoutCell(CEL.visPLOJumpDirY);
        public static Src PageShapeSplit => PageLayoutCell(CEL.visPLOSplit);
        public static Src PlaceDepth => PageLayoutCell(CEL.visPLOPlaceDepth);
        public static Src PlaceFlip => PageLayoutCell(CEL.visPLOPlaceFlip);
        public static Src PlaceStyle => PageLayoutCell(CEL.visPLOPlaceStyle);
        public static Src PlowCode => PageLayoutCell(CEL.visPLOPlowCode);
        public static Src ResizePage => PageLayoutCell(CEL.visPLOResizePage);
        public static Src RouteStyle => PageLayoutCell(CEL.visPLORouteStyle);
        public static Src AvoidPageBreaks => PageLayoutCell(CEL.visPLOAvoidPageBreaks); // new in Visio 2010

        // print properties
        private static Src PrintCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, c);
        public static Src PageLeftMargin => PrintCell(CEL.visPrintPropertiesLeftMargin);
        public static Src CenterX => PrintCell(CEL.visPrintPropertiesCenterX);
        public static Src CenterY => PrintCell(CEL.visPrintPropertiesCenterY);
        public static Src OnPage => PrintCell(CEL.visPrintPropertiesOnPage);
        public static Src PageBottomMargin => PrintCell(CEL.visPrintPropertiesBottomMargin);
        public static Src PageRightMargin => PrintCell(CEL.visPrintPropertiesRightMargin);
        public static Src PagesX => PrintCell(CEL.visPrintPropertiesPagesX);
        public static Src PagesY => PrintCell(CEL.visPrintPropertiesPagesY);
        public static Src PageTopMargin => PrintCell(CEL.visPrintPropertiesTopMargin);
        public static Src PaperKind => PrintCell(CEL.visPrintPropertiesPaperKind);
        public static Src PrintGrid => PrintCell(CEL.visPrintPropertiesPrintGrid);
        public static Src PrintPageOrientation => PrintCell(CEL.visPrintPropertiesPageOrientation);
        public static Src ScaleX => PrintCell(CEL.visPrintPropertiesScaleX);
        public static Src ScaleY => PrintCell(CEL.visPrintPropertiesScaleY);
        public static Src PaperSource => PrintCell(CEL.visPrintPropertiesPaperSource);

        // page properties
        private static Src PagePropCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPage, c);
        public static Src DrawingScale => PagePropCell(CEL.visPageDrawingScale);
        public static Src DrawingScaleType => PagePropCell(CEL.visPageDrawScaleType);
        public static Src DrawingSizeType => PagePropCell(CEL.visPageDrawSizeType);
        public static Src InhibitSnap => PagePropCell(CEL.visPageInhibitSnap);
        public static Src PageHeight => PagePropCell(CEL.visPageHeight);
        public static Src PageScale => PagePropCell(CEL.visPageScale);
        public static Src PageWidth => PagePropCell(CEL.visPageWidth);
        public static Src ShdwObliqueAngle => PagePropCell(CEL.visPageShdwObliqueAngle);
        public static Src ShdwOffsetX => PagePropCell(CEL.visPageShdwOffsetX);
        public static Src ShdwOffsetY => PagePropCell(CEL.visPageShdwOffsetY);
        public static Src ShdwScaleFactor => PagePropCell(CEL.visPageShdwScaleFactor);
        public static Src ShdwType => PagePropCell(CEL.visPageShdwType);
        public static Src UIVisibility => PagePropCell(CEL.visPageUIVisibility);
        public static Src DrawingResizeType => PagePropCell(CEL.visPageDrawResizeType); // new in Visio 2010

        // paragraph
        private static Src ParaCell(CEL c) => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, c);
        public static Src Para_Bullet => ParaCell(CEL.visBulletIndex);
        public static Src Para_BulletFont => ParaCell(CEL.visBulletFont);
        public static Src Para_BulletFontSize => ParaCell(CEL.visBulletFontSize);
        public static Src Para_BulletStr => ParaCell(CEL.visBulletString);
        public static Src Para_Flags => ParaCell(CEL.visFlags);
        public static Src Para_HorzAlign => ParaCell(CEL.visHorzAlign);
        public static Src Para_IndFirst => ParaCell(CEL.visIndentFirst);
        public static Src Para_IndLeft => ParaCell(CEL.visIndentLeft);
        public static Src Para_IndRight => ParaCell(CEL.visIndentRight);
        public static Src Para_LocalizeBulletFont => ParaCell(CEL.visLocalizeBulletFont);
        public static Src Para_SpAfter => ParaCell(CEL.visSpaceAfter);
        public static Src Para_SpBefore => ParaCell(CEL.visSpaceBefore);
        public static Src Para_SpLine => ParaCell(CEL.visSpaceLine);
        public static Src Para_TextPosAfterBullet => ParaCell(CEL.visTextPosAfterBullet);

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
        private static Src XGridCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, c);
        public static Src XGridDensity => XGridCell(CEL.visXGridDensity);
        public static Src XGridOrigin => XGridCell(CEL.visXGridOrigin);
        public static Src XGridSpacing => XGridCell(CEL.visXGridSpacing);
        public static Src XRulerDensity => XGridCell(CEL.visXRulerDensity);
        public static Src XRulerOrigin => XGridCell(CEL.visXRulerOrigin);
        public static Src YGridDensity => XGridCell(CEL.visYGridDensity);
        public static Src YGridOrigin => XGridCell(CEL.visYGridOrigin);
        public static Src YGridSpacing => XGridCell(CEL.visYGridSpacing);
        public static Src YRulerDensity => XGridCell(CEL.visYRulerDensity);
        public static Src YRulerOrigin => XGridCell(CEL.visYRulerOrigin);

        // Shape Tranform
        private static Src XFormCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXFormOut, c);
        public static Src Angle => XFormCell(CEL.visXFormAngle);
        public static Src FlipX => XFormCell(CEL.visXFormFlipX);
        public static Src FlipY => XFormCell(CEL.visXFormFlipY);
        public static Src Height => XFormCell(CEL.visXFormHeight);
        public static Src LocPinX => XFormCell(CEL.visXFormLocPinX);
        public static Src LocPinY => XFormCell(CEL.visXFormLocPinY);
        public static Src PinX => XFormCell(CEL.visXFormPinX);
        public static Src PinY => XFormCell(CEL.visXFormPinY);
        public static Src ResizeMode => XFormCell(CEL.visXFormResizeMode);
        public static Src Width => XFormCell(CEL.visXFormWidth);

        // reviewer
        private static Src ReviewerCell(CEL c) => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, c);
        public static Src Reviewer_Color => ReviewerCell(CEL.visReviewerColor);
        public static Src Reviewer_Initials => ReviewerCell(CEL.visReviewerInitials);
        public static Src Reviewer_Name => ReviewerCell(CEL.visReviewerName);

        // shape data
        private static Src PropCell(CEL c) => new Src(SEC.visSectionProp, ROW.visRowProp, c);
        public static Src Prop_SortKey => PropCell(CEL.visCustPropsSortKey);
        public static Src Prop_Ask => PropCell(CEL.visCustPropsAsk);
        public static Src Prop_Calendar => PropCell(CEL.visCustPropsCalendar);
        public static Src Prop_Format => PropCell(CEL.visCustPropsFormat);
        public static Src Prop_Invisible => PropCell(CEL.visCustPropsInvis);
        public static Src Prop_Label => PropCell(CEL.visCustPropsLabel);
        public static Src Prop_LangID => PropCell(CEL.visCustPropsLangID);
        public static Src Prop_Prompt => PropCell(CEL.visCustPropsPrompt);
        public static Src Prop_Type => PropCell(CEL.visCustPropsType);
        public static Src Prop_Value => PropCell(CEL.visCustPropsValue);

        // Layers
        private static Src LayerCell(CEL c) => new Src(SEC.visSectionLayer, ROW.visRowLayer, c);
        public static Src Layers_Active => LayerCell(CEL.visLayerActive);
        public static Src Layers_Color => LayerCell(CEL.visLayerColor);
        public static Src Layers_Glue => LayerCell(CEL.visLayerGlue);
        public static Src Layers_Locked => LayerCell(CEL.visLayerLock);
        public static Src Layers_Print => LayerCell(CEL.visDocPreviewScope);
        public static Src Layers_Snap => LayerCell(CEL.visLayerSnap);
        public static Src Layers_ColorTrans => LayerCell(CEL.visLayerColorTrans);
        public static Src Layers_Visible => LayerCell(CEL.visLayerVisible);

        //text transform
        private static Src TextXFormCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowTextXForm, c);
        public static Src TxtAngle => TextXFormCell(CEL.visXFormAngle);
        public static Src TxtHeight => TextXFormCell(CEL.visXFormHeight);
        public static Src TxtLocPinX => TextXFormCell(CEL.visXFormLocPinX);
        public static Src TxtLocPinY => TextXFormCell(CEL.visXFormLocPinY);
        public static Src TxtPinX => TextXFormCell(CEL.visXFormPinX);
        public static Src TxtPinY => TextXFormCell(CEL.visXFormPinY);
        public static Src TxtWidth => TextXFormCell(CEL.visXFormWidth);

        // user defined cells
        private static Src UserDefCell(CEL c) => new Src(SEC.visSectionUser, ROW.visRowUser, c);
        public static Src User_Prompt => UserDefCell(CEL.visUserPrompt);
        public static Src User_Value => UserDefCell(CEL.visUserValue);

        // Fields
        private static Src FieldCell(CEL c) => new Src(SEC.visSectionTextField, ROW.visRowField, c);
        public static Src Fields_Calendar => FieldCell(CEL.visFieldCalendar);
        public static Src Fields_Format => FieldCell(CEL.visFieldFormat);
        public static Src Fields_ObjectKind => FieldCell(CEL.visFieldObjectKind);
        public static Src Fields_Type => FieldCell(CEL.visFieldType);
        public static Src Fields_UICat => FieldCell(CEL.visFieldUICategory);
        public static Src Fields_UICod => FieldCell(CEL.visFieldUICode);
        public static Src Fields_UIFmt => FieldCell(CEL.visFieldUIFormat);
        public static Src Fields_Value => FieldCell(CEL.visFieldCell);

        // text block format
        private static Src TextBlockCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowText, c);
        public static Src BottomMargin => TextBlockCell(CEL.visTxtBlkBottomMargin);
        public static Src DefaultTabStop => TextBlockCell(CEL.visTxtBlkDefaultTabStop);
        public static Src LeftMargin => TextBlockCell(CEL.visTxtBlkLeftMargin);
        public static Src RightMargin => TextBlockCell(CEL.visTxtBlkRightMargin);
        public static Src TextBkgnd => TextBlockCell(CEL.visTxtBlkBkgnd);
        public static Src TextBkgndTrans => TextBlockCell(CEL.visTxtBlkBkgndTrans);
        public static Src TextDirection => TextBlockCell(CEL.visTxtBlkDirection);
        public static Src TopMargin => TextBlockCell(CEL.visTxtBlkTopMargin);
        public static Src VerticalAlign => TextBlockCell(CEL.visTxtBlkVerticalAlign);

        // Action tags
        private static Src SmartTagCell(CEL c) => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, c);
        public static Src SmartTags_ButtonFace => SmartTagCell(CEL.visSmartTagButtonFace);
        public static Src SmartTags_Description => SmartTagCell(CEL.visSmartTagDescription);
        public static Src SmartTags_Disabled => SmartTagCell(CEL.visSmartTagDisabled);
        public static Src SmartTags_DisplayMode => SmartTagCell(CEL.visSmartTagDisplayMode);
        public static Src SmartTags_TagName => SmartTagCell(CEL.visSmartTagName);
        public static Src SmartTags_X => SmartTagCell(CEL.visSmartTagX);
        public static Src SmartTags_XJustify => SmartTagCell(CEL.visSmartTagXJustify);
        public static Src SmartTags_Y => SmartTagCell(CEL.visSmartTagY);
        public static Src SmartTags_YJustify => SmartTagCell(CEL.visSmartTagYJustify);

        // style
        private static Src StyleCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowStyle, c);
        public static Src EnableFillProps => StyleCell(CEL.visStyleIncludesFill);
        public static Src EnableLineProps => StyleCell(CEL.visStyleIncludesLine);
        public static Src EnableTextProps => StyleCell(CEL.visStyleIncludesText);

        //tabs
        private static Src TabCell(CEL c) => new Src(SEC.visSectionTab, ROW.visRowTab, c);
        public static Src Tabs_Alignment => TabCell(CEL.visTabAlign);
        public static Src Tabs_Position => TabCell(CEL.visTabPos);
        public static Src Tabs_StopCount => TabCell(CEL.visTabStopCount);

        // shape layout
        private static Src ShapeLayoutCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, c);
        public static Src ConFixedCode => ShapeLayoutCell(CEL.visSLOConFixedCode);
        public static Src ConLineJumpCode => ShapeLayoutCell(CEL.visSLOJumpCode);
        public static Src ConLineJumpDirX => ShapeLayoutCell(CEL.visSLOJumpDirX);
        public static Src ConLineJumpDirY => ShapeLayoutCell(CEL.visSLOJumpDirY);
        public static Src ConLineJumpStyle => ShapeLayoutCell(CEL.visSLOJumpStyle);
        public static Src ConLineRouteExt => ShapeLayoutCell(CEL.visSLOLineRouteExt);
        public static Src ShapeFixedCode => ShapeLayoutCell(CEL.visSLOFixedCode);
        public static Src ShapePermeablePlace => ShapeLayoutCell(CEL.visSLOPermeablePlace);
        public static Src ShapePermeableX => ShapeLayoutCell(CEL.visSLOPermX);
        public static Src ShapePermeableY => ShapeLayoutCell(CEL.visSLOPermY);
        public static Src ShapePlaceFlip => ShapeLayoutCell(CEL.visSLOPlaceFlip);
        public static Src ShapePlaceStyle => ShapeLayoutCell(CEL.visSLOPlaceStyle);
        public static Src ShapePlowCode => ShapeLayoutCell(CEL.visSLOPlowCode);
        public static Src ShapeRouteStyle => ShapeLayoutCell(CEL.visSLORouteStyle);
        public static Src ShapeSplit => ShapeLayoutCell(CEL.visSLOSplit);
        public static Src ShapeSplittable => ShapeLayoutCell(CEL.visSLOSplittable);
        public static Src DisplayLevel => ShapeLayoutCell(CEL.visSLODisplayLevel); // new in Visio 2010
        public static Src Relationships => ShapeLayoutCell(CEL.visSLORelationships); // new in Visio 2010
    }

}
