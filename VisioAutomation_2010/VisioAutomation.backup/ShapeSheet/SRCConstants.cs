using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;

namespace VisioAutomation.ShapeSheet
{
    public static class SrcConstants
    {
        // Actions
        private static Src _action_cell(CEL c) => new Src(SEC.visSectionAction, ROW.visRowAction, c);
        public static Src ActionAction => _action_cell(CEL.visActionAction);
        public static Src ActionBeginGroup => _action_cell(CEL.visActionBeginGroup);
        public static Src ActionButtonFace => _action_cell(CEL.visActionButtonFace);
        public static Src ActionChecked => _action_cell(CEL.visActionChecked);
        public static Src ActionDisabled => _action_cell(CEL.visActionDisabled);
        public static Src ActionInvisible => _action_cell(CEL.visActionInvisible);
        public static Src ActionMenu => _action_cell(CEL.visActionMenu);
        public static Src ActionReadOnly => _action_cell(CEL.visActionReadOnly);
        public static Src ActionSortKey => _action_cell(CEL.visActionSortKey);
        public static Src ActionTagName => _action_cell(CEL.visActionTagName);
        public static Src ActionFlyoutChild => _action_cell(CEL.visActionFlyoutChild); // new for visio 2010

        // Alignment
        private static Src _align_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowAlign, c);
        public static Src AlignBottom => _align_cell(CEL.visAlignBottom);
        public static Src AlignCenter => _align_cell(CEL.visAlignCenter);
        public static Src AlignLeft => _align_cell(CEL.visAlignLeft);
        public static Src AlignMiddle => _align_cell(CEL.visAlignMiddle);
        public static Src AlignRight => _align_cell(CEL.visAlignRight);
        public static Src AlignTop => _align_cell(CEL.visAlignTop);

        // Annotation
        private static Src _annotation_cell(CEL c) => new Src(SEC.visSectionAnnotation, ROW.visRowAnnotation, c);
        public static Src AnnotationComment => _annotation_cell(CEL.visAnnotationComment);
        public static Src AnnotationDate => _annotation_cell(CEL.visAnnotationDate);
        public static Src AnnotationLangID => _annotation_cell(CEL.visAnnotationLangID);
        public static Src AnnotationMarkerIndex => _annotation_cell(CEL.visAnnotationMarkerIndex);
        public static Src AnnotationX => _annotation_cell(CEL.visAnnotationX);
        public static Src AnnotationY => _annotation_cell(CEL.visAnnotationY);

        // Character
        private static Src _char_cell(CEL c) => new Src(SEC.visSectionCharacter, ROW.visRowCharacter, c);
        public static Src CharAsianFont => _char_cell(CEL.visCharacterAsianFont);
        public static Src CharCase => _char_cell(CEL.visCharacterCase);
        public static Src CharColor => _char_cell(CEL.visCharacterColor);
        public static Src CharComplexScriptFont => _char_cell(CEL.visCharacterComplexScriptFont);
        public static Src CharComplexScriptSize => _char_cell(CEL.visCharacterComplexScriptSize);
        public static Src CharDoubleStrikethrough => _char_cell(CEL.visCharacterDoubleStrikethrough);
        public static Src CharDoubleUnderline => _char_cell(CEL.visCharacterDblUnderline);
        public static Src CharFont => _char_cell(CEL.visCharacterFont);
        public static Src CharLangID => _char_cell(CEL.visCharacterLangID);
        public static Src CharLocale => _char_cell(CEL.visCharacterLocale);
        public static Src CharLocalizeFont => _char_cell(CEL.visCharacterLocalizeFont);
        public static Src CharOverline => _char_cell(CEL.visCharacterOverline);
        public static Src CharPerpendicular => _char_cell(CEL.visCharacterPerpendicular);
        public static Src CharPos => _char_cell(CEL.visCharacterPos);
        public static Src CharRTLText => _char_cell(CEL.visCharacterRTLText);
        public static Src CharFontScale => _char_cell(CEL.visCharacterFontScale);
        public static Src CharSize => _char_cell(CEL.visCharacterSize);
        public static Src CharLetterspace => _char_cell(CEL.visCharacterLetterspace);
        public static Src CharStrikethru => _char_cell(CEL.visCharacterStrikethru);
        public static Src CharStyle => _char_cell(CEL.visCharacterStyle);
        public static Src CharColorTransparency => _char_cell(CEL.visCharacterColorTrans);
        public static Src CharUseVertical => _char_cell(CEL.visCharacterUseVertical);

        // Connections
        private static Src _connection_point_cell(CEL c) => new Src(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, c);
        public static Src ConnectionPointD => _connection_point_cell(CEL.visCnnctD);
        public static Src ConnectionPointDirX => _connection_point_cell(CEL.visCnnctDirX);
        public static Src ConnectionPointDirY => _connection_point_cell(CEL.visCnnctDirY);
        public static Src ConnectionPointType => _connection_point_cell(CEL.visCnnctType);
        public static Src ConnectionPointX => _connection_point_cell(CEL.visX);
        public static Src ConnectionPointY => _connection_point_cell(CEL.visY);

        // Controls
        private static Src _control_cell(CEL c) => new Src(SEC.visSectionControls, ROW.visRowControl, c);
        public static Src ControlCanGlue => _control_cell(CEL.visCtlGlue);
        public static Src ControlTip => _control_cell(CEL.visCtlTip);
        public static Src ControlXBehavior => _control_cell(CEL.visCtlXCon);
        public static Src ControlX => _control_cell(CEL.visCtlX);
        public static Src ControlXDynamics => _control_cell(CEL.visCtlXDyn);
        public static Src ControlYBehavior => _control_cell(CEL.visCtlYCon);
        public static Src ControlY => _control_cell(CEL.visCtlY);
        public static Src ControlYDynamics => _control_cell(CEL.visCtlYDyn);

        // Document Properties
        private static Src _doc_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowDoc, c);
        public static Src DocAddMarkup => _doc_cell(CEL.visDocAddMarkup);
        public static Src DocLangID => _doc_cell(CEL.visDocLangID);
        public static Src DocLockPreview => _doc_cell(CEL.visDocLockPreview);
        public static Src DocOutputFormat => _doc_cell(CEL.visDocOutputFormat);
        public static Src DocPreviewQuality => _doc_cell(CEL.visDocPreviewQuality);
        public static Src DocPreviewScope => _doc_cell(CEL.visDocPreviewScope);
        public static Src DocViewMarkup => _doc_cell(CEL.visDocViewMarkup);

        // Events
        private static Src _event_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowEvent, c);
        public static Src EventDoubleClick => _event_cell(CEL.visEvtCellDblClick);
        public static Src EventDrop => _event_cell(CEL.visEvtCellDrop);
        public static Src EventMultiDrop => _event_cell(CEL.visEvtCellMultiDrop);
        public static Src EventXFormMod => _event_cell(CEL.visEvtCellXFMod);
        public static Src EventTheText => _event_cell(CEL.visEvtCellTheText);

        // ForeignImageInfo
        private static Src _foreign_img_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowForeign, c);
        public static Src ForeignImageHeight => _foreign_img_cell(CEL.visFrgnImgHeight);
        public static Src ForeignImageOffsetX => _foreign_img_cell(CEL.visFrgnImgOffsetX);
        public static Src ForeignImageOffsetY => _foreign_img_cell(CEL.visFrgnImgOffsetY);
        public static Src ForeignImageWidth => _foreign_img_cell(CEL.visFrgnImgWidth);

        // Geometry 
        private static Src _geometry_vertex_cell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowVertex, c);
        public static Src GeometryVertexA => _geometry_vertex_cell(CEL.visBow);
        public static Src GeometryVertexB => _geometry_vertex_cell(CEL.visControlX);
        public static Src GeometryVertexC => _geometry_vertex_cell(CEL.visEccentricityAngle);
        public static Src GeometryVertexD => _geometry_vertex_cell(CEL.visAspectRatio);
        public static Src GeometryVertexE => _geometry_vertex_cell(CEL.visNURBSData);
        public static Src GeometryVertexX => _geometry_vertex_cell(CEL.visX);
        public static Src GeometryVertexY => _geometry_vertex_cell(CEL.visY);

        // Geometry
        private static Src _geometry_row_cell(CEL c) => new Src(SEC.visSectionFirstComponent, ROW.visRowComponent, c);
        public static Src GeometryNoFill => _geometry_row_cell(CEL.visCompNoFill);
        public static Src GeometryNoLine => _geometry_row_cell(CEL.visCompNoLine);
        public static Src GeometryNoShow => _geometry_row_cell(CEL.visCompNoShow);
        public static Src GeometryNoSnap => _geometry_row_cell(CEL.visCompNoSnap);
        public static Src GeometryNoQuickDrag => _geometry_row_cell(CEL.visCompNoQuickDrag);

        // Fill Format
        private static Src _fill_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowFill, c);
        public static Src FillBackground => _fill_cell(CEL.visFillBkgnd);
        public static Src FillBackgroundTransparency => _fill_cell(CEL.visFillBkgndTrans);
        public static Src FillForeground => _fill_cell(CEL.visFillForegnd);
        public static Src FillForegroundTransparency => _fill_cell(CEL.visFillForegndTrans);
        public static Src FillPattern => _fill_cell(CEL.visFillPattern);
        public static Src FillShadowObliqueAngle => _fill_cell(CEL.visFillShdwObliqueAngle);
        public static Src FillShadowOffsetX => _fill_cell(CEL.visFillShdwOffsetX);
        public static Src FillShadowOffsetY => _fill_cell(CEL.visFillShdwOffsetY);
        public static Src FillShadowScaleFactor => _fill_cell(CEL.visFillShdwScaleFactor);
        public static Src FillShadowType => _fill_cell(CEL.visFillShdwType);
        public static Src FillShadowBackground => _fill_cell(CEL.visFillShdwBkgnd);
        public static Src FillShadowBackgroundTransparency => _fill_cell(CEL.visFillShdwBkgndTrans);
        public static Src FillShadowForeground => _fill_cell(CEL.visFillShdwForegnd);
        public static Src FillShadowForegroundTransparency => _fill_cell(CEL.visFillShdwForegndTrans);
        public static Src FillShadowPattern => _fill_cell(CEL.visFillShdwPattern);

        // GlueInfo
        private static Src GlueCell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src GlueBeginTrigger => GlueCell(CEL.visBegTrigger);
        public static Src GlueEndTrigger => GlueCell(CEL.visEndTrigger);
        public static Src GlueType => GlueCell(CEL.visGlueType);
        public static Src GlueWalkPref => GlueCell(CEL.visWalkPref);

        // GroupProperties
        private static Src _group_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowGroup, c);
        public static Src GroupDisplayMode => _group_cell(CEL.visGroupDisplayMode);
        public static Src GroupDontMoveChildren => _group_cell(CEL.visGroupDontMoveChildren);
        public static Src GroupIsDropTarget => _group_cell(CEL.visGroupIsDropTarget);
        public static Src GroupIsSnapTarget => _group_cell(CEL.visGroupIsSnapTarget);
        public static Src GroupIsTextEditTarget => _group_cell(CEL.visGroupIsTextEditTarget);
        public static Src GroupSelectMode => _group_cell(CEL.visGroupSelectMode);

        // Hyperlinks
        private static Src _hyperlink_cell(CEL c) => new Src(SEC.visSectionHyperlink, ROW.visRowHyperlink, c);
        public static Src HyperlinkAddress => _hyperlink_cell(CEL.visHLinkAddress);
        public static Src HyperlinkDefault => _hyperlink_cell(CEL.visHLinkDefault);
        public static Src HyperlinkDescription => _hyperlink_cell(CEL.visHLinkDescription);
        public static Src HyperlinkExtraInfo => _hyperlink_cell(CEL.visHLinkExtraInfo);
        public static Src HyperlinkFrame => _hyperlink_cell(CEL.visHLinkFrame);
        public static Src HyperlinkInvisible => _hyperlink_cell(CEL.visHLinkInvisible);
        public static Src HyperlinkNewWindow => _hyperlink_cell(CEL.visHLinkNewWin);
        public static Src HyperlinkSortKey => _hyperlink_cell(CEL.visHLinkSortKey);
        public static Src HyperlinkSubAddress => _hyperlink_cell(CEL.visHLinkSubAddress);

        // Image Properties
        private static Src _image_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowImage, c);
        public static Src ImageBlur => _image_cell(CEL.visImageBlur);
        public static Src ImageBrightness => _image_cell(CEL.visImageBrightness);
        public static Src ImageContrast => _image_cell(CEL.visImageContrast);
        public static Src ImageDenoise => _image_cell(CEL.visImageDenoise);
        public static Src ImageGamma => _image_cell(CEL.visImageGamma);
        public static Src ImageSharpen => _image_cell(CEL.visImageSharpen);
        public static Src ImageTransparency => _image_cell(CEL.visImageTransparency);

        // Line format
        private static Src _line_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowLine, c);
        public static Src LineBeginArrow => _line_cell(CEL.visLineBeginArrow);
        public static Src LineBeginArrowSize => _line_cell(CEL.visLineBeginArrowSize);
        public static Src LineEndArrow => _line_cell(CEL.visLineEndArrow);
        public static Src LineEndArrowSize => _line_cell(CEL.visLineEndArrowSize);
        public static Src LineCap => _line_cell(CEL.visLineEndCap);
        public static Src LineColor => _line_cell(CEL.visLineColor);
        public static Src LineColorTransparency => _line_cell(CEL.visLineColorTrans);
        public static Src LinePattern => _line_cell(CEL.visLinePattern);
        public static Src LineWeight => _line_cell(CEL.visLineWeight);
        public static Src LineRounding => _line_cell(CEL.visLineRounding);

        // Miscellaneous
        private static Src _misc_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowMisc, c);
        public static Src MiscCalendar => _misc_cell(CEL.visObjCalendar);
        public static Src MiscComment => _misc_cell(CEL.visComment);
        public static Src MiscDropOnPageScale => _misc_cell(CEL.visObjDropOnPageScale);
        public static Src MiscDynFeedback => _misc_cell(CEL.visDynFeedback);
        public static Src MiscIsDropSource => _misc_cell(CEL.visDropSource);
        public static Src MiscLangID => _misc_cell(CEL.visObjLangID);
        public static Src MiscLocalizeMerge => _misc_cell(CEL.visObjLocalizeMerge);
        public static Src MiscNoAlignBox => _misc_cell(CEL.visNoAlignBox);
        public static Src MiscNoControlHandles => _misc_cell(CEL.visNoCtlHandles);
        public static Src MiscNoLiveDynamics => _misc_cell(CEL.visNoLiveDynamics);
        public static Src MiscNonPrinting => _misc_cell(CEL.visNonPrinting);
        public static Src MiscNoObjHandles => _misc_cell(CEL.visNoObjHandles);
        public static Src MiscObjType => _misc_cell(CEL.visLOFlags);
        public static Src MiscUpdateAlignBox => _misc_cell(CEL.visUpdateAlignBox);
        public static Src MiscHideText => _misc_cell(CEL.visHideText);

        // 1d endpoints
        private static Src _one_d_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXForm1D, c);
        public static Src OneDBeginX => _one_d_cell(CEL.vis1DBeginX);
        public static Src OneDBeginY => _one_d_cell(CEL.vis1DBeginY);
        public static Src OneDEndX => _one_d_cell(CEL.vis1DEndX);
        public static Src OneDEndY => _one_d_cell(CEL.vis1DEndY);

        // page layout
        private static Src _page_layout_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPageLayout, c);
        public static Src PageLayoutAvenueSizeX => _page_layout_cell(CEL.visPLOAvenueSizeX);
        public static Src PageLayoutAvenueSizeY => _page_layout_cell(CEL.visPLOAvenueSizeY);
        public static Src PageLayoutBlockSizeX => _page_layout_cell(CEL.visPLOBlockSizeX);
        public static Src PageLayoutBlockSizeY => _page_layout_cell(CEL.visPLOBlockSizeY);
        public static Src PageLayoutControlAsInput => _page_layout_cell(CEL.visPLOCtrlAsInput);
        public static Src PageLayoutDynamicsOff => _page_layout_cell(CEL.visPLODynamicsOff);
        public static Src PageLayoutEnableGrid => _page_layout_cell(CEL.visPLOEnableGrid);
        public static Src PageLayoutLineAdjustFrom => _page_layout_cell(CEL.visPLOLineAdjustFrom);
        public static Src PageLayoutLineAdjustTo => _page_layout_cell(CEL.visPLOLineAdjustTo);
        public static Src PageLayoutLineJumpCode => _page_layout_cell(CEL.visPLOJumpCode);
        public static Src PageLayoutLineJumpFactorX => _page_layout_cell(CEL.visPLOJumpFactorX);
        public static Src PageLayoutLineJumpFactorY => _page_layout_cell(CEL.visPLOJumpFactorY);
        public static Src PageLayoutLineJumpStyle => _page_layout_cell(CEL.visPLOJumpStyle);
        public static Src PageLayoutLineRouteExt => _page_layout_cell(CEL.visPLOLineRouteExt);
        public static Src PageLayoutLineToLineX => _page_layout_cell(CEL.visPLOLineToLineX);
        public static Src PageLayoutLineToLineY => _page_layout_cell(CEL.visPLOLineToLineY);
        public static Src PageLayoutLineToNodeX => _page_layout_cell(CEL.visPLOLineToNodeX);
        public static Src PageLayoutLineToNodeY => _page_layout_cell(CEL.visPLOLineToNodeY);
        public static Src PageLayoutLineJumpDirX => _page_layout_cell(CEL.visPLOJumpDirX);
        public static Src PageLayoutLineJumpDirY => _page_layout_cell(CEL.visPLOJumpDirY);
        public static Src PageLayoutShapeSplit => _page_layout_cell(CEL.visPLOSplit);
        public static Src PageLayoutPlaceDepth => _page_layout_cell(CEL.visPLOPlaceDepth);
        public static Src PageLayoutPlaceFlip => _page_layout_cell(CEL.visPLOPlaceFlip);
        public static Src PageLayoutPlaceStyle => _page_layout_cell(CEL.visPLOPlaceStyle);
        public static Src PageLayoutPlowCode => _page_layout_cell(CEL.visPLOPlowCode);
        public static Src PageLayoutResizePage => _page_layout_cell(CEL.visPLOResizePage);
        public static Src PageLayoutRouteStyle => _page_layout_cell(CEL.visPLORouteStyle);
        public static Src PageLayoutAvoidPageBreaks => _page_layout_cell(CEL.visPLOAvoidPageBreaks); // new in Visio 2010

        // print properties
        private static Src _print_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPrintProperties, c);
        public static Src PrintLeftMargin => _print_cell(CEL.visPrintPropertiesLeftMargin);
        public static Src PrintCenterX => _print_cell(CEL.visPrintPropertiesCenterX);
        public static Src PrintCenterY => _print_cell(CEL.visPrintPropertiesCenterY);
        public static Src PrintOnPage => _print_cell(CEL.visPrintPropertiesOnPage);
        public static Src PrintBottomMargin => _print_cell(CEL.visPrintPropertiesBottomMargin);
        public static Src PrintRightMargin => _print_cell(CEL.visPrintPropertiesRightMargin);
        public static Src PrintPagesX => _print_cell(CEL.visPrintPropertiesPagesX);
        public static Src PrintPagesY => _print_cell(CEL.visPrintPropertiesPagesY);
        public static Src PrintTopMargin => _print_cell(CEL.visPrintPropertiesTopMargin);
        public static Src PrintPaperKind => _print_cell(CEL.visPrintPropertiesPaperKind);
        public static Src PrintGrid => _print_cell(CEL.visPrintPropertiesPrintGrid);
        public static Src PrintPageOrientation => _print_cell(CEL.visPrintPropertiesPageOrientation);
        public static Src PrintScaleX => _print_cell(CEL.visPrintPropertiesScaleX);
        public static Src PrintScaleY => _print_cell(CEL.visPrintPropertiesScaleY);
        public static Src PrintPaperSource => _print_cell(CEL.visPrintPropertiesPaperSource);

        // page properties
        private static Src _page_prop_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowPage, c);
        public static Src PageDrawingScale => _page_prop_cell(CEL.visPageDrawingScale);
        public static Src PageDrawingScaleType => _page_prop_cell(CEL.visPageDrawScaleType);
        public static Src PageDrawingSizeType => _page_prop_cell(CEL.visPageDrawSizeType);
        public static Src PageInhibitSnap => _page_prop_cell(CEL.visPageInhibitSnap);
        public static Src PageHeight => _page_prop_cell(CEL.visPageHeight);
        public static Src PageScale => _page_prop_cell(CEL.visPageScale);
        public static Src PageWidth => _page_prop_cell(CEL.visPageWidth);
        public static Src PageShadowObliqueAngle => _page_prop_cell(CEL.visPageShdwObliqueAngle);
        public static Src PageShadowOffsetX => _page_prop_cell(CEL.visPageShdwOffsetX);
        public static Src PageShadowOffsetY => _page_prop_cell(CEL.visPageShdwOffsetY);
        public static Src PageShadowScaleFactor => _page_prop_cell(CEL.visPageShdwScaleFactor);
        public static Src PageShadowType => _page_prop_cell(CEL.visPageShdwType);
        public static Src PageUIVisibility => _page_prop_cell(CEL.visPageUIVisibility);
        public static Src PageDrawingResizeType => _page_prop_cell(CEL.visPageDrawResizeType); // new in Visio 2010

        // paragraph
        private static Src _para_cell(CEL c) => new Src(SEC.visSectionParagraph, ROW.visRowParagraph, c);
        public static Src ParaBullet => _para_cell(CEL.visBulletIndex);
        public static Src ParaBulletFont => _para_cell(CEL.visBulletFont);
        public static Src ParaBulletFontSize => _para_cell(CEL.visBulletFontSize);
        public static Src ParaBulletString => _para_cell(CEL.visBulletString);
        public static Src ParaFlags => _para_cell(CEL.visFlags);
        public static Src ParaHorizontalAlign => _para_cell(CEL.visHorzAlign);
        public static Src ParaIndentFirst => _para_cell(CEL.visIndentFirst);
        public static Src ParaIndentLeft => _para_cell(CEL.visIndentLeft);
        public static Src ParaIndentRight => _para_cell(CEL.visIndentRight);
        public static Src ParaLocalizeBulletFont => _para_cell(CEL.visLocalizeBulletFont);
        public static Src ParaSpacingAfter => _para_cell(CEL.visSpaceAfter);
        public static Src ParaSpacingBefore => _para_cell(CEL.visSpaceBefore);
        public static Src ParaSpacingLine => _para_cell(CEL.visSpaceLine);
        public static Src ParaTextPosAfterBullet => _para_cell(CEL.visTextPosAfterBullet);

        // protection
        private static Src _lock_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowLock, c);
        public static Src LockAspect => _lock_cell(CEL.visLockAspect);
        public static Src LockBegin => _lock_cell(CEL.visLockBegin);
        public static Src LockCalcWH => _lock_cell(CEL.visLockCalcWH);
        public static Src LockCrop => _lock_cell(CEL.visLockCrop);
        public static Src LockCustomProp => _lock_cell(CEL.visLockCustProp);
        public static Src LockDelete => _lock_cell(CEL.visLockDelete);
        public static Src LockEnd => _lock_cell(CEL.visLockEnd);
        public static Src LockFormat => _lock_cell(CEL.visLockFormat);
        public static Src LockFromGroupFormat => _lock_cell(CEL.visLockFromGroupFormat);
        public static Src LockGroup => _lock_cell(CEL.visLockGroup);
        public static Src LockHeight => _lock_cell(CEL.visLockHeight);
        public static Src LockMoveX => _lock_cell(CEL.visLockMoveX);
        public static Src LockMoveY => _lock_cell(CEL.visLockMoveY);
        public static Src LockRotate => _lock_cell(CEL.visLockRotate);
        public static Src LockSelect => _lock_cell(CEL.visLockSelect);
        public static Src LockTextEdit => _lock_cell(CEL.visLockTextEdit);
        public static Src LockThemeColors => _lock_cell(CEL.visLockThemeColors);
        public static Src LockThemeEffects => _lock_cell(CEL.visLockThemeEffects);
        public static Src LockVertexEdit => _lock_cell(CEL.visLockVtxEdit);
        public static Src LockWidth => _lock_cell(CEL.visLockWidth);

        // ruler and grid
        private static Src _ruler_grid_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowRulerGrid, c);
        public static Src XGridDensity => _ruler_grid_cell(CEL.visXGridDensity);
        public static Src XGridOrigin => _ruler_grid_cell(CEL.visXGridOrigin);
        public static Src XGridSpacing => _ruler_grid_cell(CEL.visXGridSpacing);
        public static Src YGridDensity => _ruler_grid_cell(CEL.visYGridDensity);
        public static Src YGridOrigin => _ruler_grid_cell(CEL.visYGridOrigin);
        public static Src YGridSpacing => _ruler_grid_cell(CEL.visYGridSpacing);
        public static Src XRulerDensity => _ruler_grid_cell(CEL.visXRulerDensity);
        public static Src XRulerOrigin => _ruler_grid_cell(CEL.visXRulerOrigin);
        public static Src YRulerDensity => _ruler_grid_cell(CEL.visYRulerDensity);
        public static Src YRulerOrigin => _ruler_grid_cell(CEL.visYRulerOrigin);

        // Shape Tranform
        private static Src _x_form_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowXFormOut, c);
        public static Src XFormAngle => _x_form_cell(CEL.visXFormAngle);
        public static Src XFormFlipX => _x_form_cell(CEL.visXFormFlipX);
        public static Src XFormFlipY => _x_form_cell(CEL.visXFormFlipY);
        public static Src XFormHeight => _x_form_cell(CEL.visXFormHeight);
        public static Src XFormLocPinX => _x_form_cell(CEL.visXFormLocPinX);
        public static Src XFormLocPinY => _x_form_cell(CEL.visXFormLocPinY);
        public static Src XFormPinX => _x_form_cell(CEL.visXFormPinX);
        public static Src XFormPinY => _x_form_cell(CEL.visXFormPinY);
        public static Src XFormResizeMode => _x_form_cell(CEL.visXFormResizeMode);
        public static Src XFormWidth => _x_form_cell(CEL.visXFormWidth);

        // reviewer
        private static Src _reviewer_cell(CEL c) => new Src(SEC.visSectionReviewer, ROW.visRowReviewer, c);
        public static Src ReviewerColor => _reviewer_cell(CEL.visReviewerColor);
        public static Src ReviewerInitials => _reviewer_cell(CEL.visReviewerInitials);
        public static Src ReviewerName => _reviewer_cell(CEL.visReviewerName);

        // shape data
        private static Src _custom_prop_cell(CEL c) => new Src(SEC.visSectionProp, ROW.visRowProp, c);
        public static Src CustomPropSortKey => _custom_prop_cell(CEL.visCustPropsSortKey);
        public static Src CustomPropAsk => _custom_prop_cell(CEL.visCustPropsAsk);
        public static Src CustomPropCalendar => _custom_prop_cell(CEL.visCustPropsCalendar);
        public static Src CustomPropFormat => _custom_prop_cell(CEL.visCustPropsFormat);
        public static Src CustomPropInvisible => _custom_prop_cell(CEL.visCustPropsInvis);
        public static Src CustomPropLabel => _custom_prop_cell(CEL.visCustPropsLabel);
        public static Src CustomPropLangID => _custom_prop_cell(CEL.visCustPropsLangID);
        public static Src CustomPropPrompt => _custom_prop_cell(CEL.visCustPropsPrompt);
        public static Src CustomPropType => _custom_prop_cell(CEL.visCustPropsType);
        public static Src CustomPropValue => _custom_prop_cell(CEL.visCustPropsValue);

        // Layers
        private static Src _layer_cell(CEL c) => new Src(SEC.visSectionLayer, ROW.visRowLayer, c);
        public static Src LayerActive => _layer_cell(CEL.visLayerActive);
        public static Src LayerColor => _layer_cell(CEL.visLayerColor);
        public static Src LayerGlue => _layer_cell(CEL.visLayerGlue);
        public static Src LayerLocked => _layer_cell(CEL.visLayerLock);
        public static Src LayerPrint => _layer_cell(CEL.visDocPreviewScope);
        public static Src LayerSnap => _layer_cell(CEL.visLayerSnap);
        public static Src LayerColorTransparency => _layer_cell(CEL.visLayerColorTrans);
        public static Src LayerVisible => _layer_cell(CEL.visLayerVisible);

        //text transform
        private static Src _text_x_form_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowTextXForm, c);
        public static Src TextXFormAngle => _text_x_form_cell(CEL.visXFormAngle);
        public static Src TextXFormHeight => _text_x_form_cell(CEL.visXFormHeight);
        public static Src TextXFormLocPinX => _text_x_form_cell(CEL.visXFormLocPinX);
        public static Src TextXFormLocPinY => _text_x_form_cell(CEL.visXFormLocPinY);
        public static Src TextXFormPinX => _text_x_form_cell(CEL.visXFormPinX);
        public static Src TextXFormPinY => _text_x_form_cell(CEL.visXFormPinY);
        public static Src TextXFormWidth => _text_x_form_cell(CEL.visXFormWidth);

        // user defined cells
        private static Src _user_def_cell(CEL c) => new Src(SEC.visSectionUser, ROW.visRowUser, c);
        public static Src UserDefCellPrompt => _user_def_cell(CEL.visUserPrompt);
        public static Src UserDefCellValue => _user_def_cell(CEL.visUserValue);

        // Fields
        private static Src _field_cell(CEL c) => new Src(SEC.visSectionTextField, ROW.visRowField, c);
        public static Src FieldCalendar => _field_cell(CEL.visFieldCalendar);
        public static Src FieldFormat => _field_cell(CEL.visFieldFormat);
        public static Src FieldObjectKind => _field_cell(CEL.visFieldObjectKind);
        public static Src FieldType => _field_cell(CEL.visFieldType);
        public static Src FieldUICategory => _field_cell(CEL.visFieldUICategory);
        public static Src FieldUICode => _field_cell(CEL.visFieldUICode);
        public static Src FieldUIFormat => _field_cell(CEL.visFieldUIFormat);
        public static Src FieldValue => _field_cell(CEL.visFieldCell);

        // text block format
        private static Src _text_block_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowText, c);
        public static Src TextBlockBottomMargin => _text_block_cell(CEL.visTxtBlkBottomMargin);
        public static Src TextBlockDefaultTabStop => _text_block_cell(CEL.visTxtBlkDefaultTabStop);
        public static Src TextBlockLeftMargin => _text_block_cell(CEL.visTxtBlkLeftMargin);
        public static Src TextBlockRightMargin => _text_block_cell(CEL.visTxtBlkRightMargin);
        public static Src TextBlockBackground => _text_block_cell(CEL.visTxtBlkBkgnd);
        public static Src TextBlockBackgroundTransparency => _text_block_cell(CEL.visTxtBlkBkgndTrans);
        public static Src TextBlockDirection => _text_block_cell(CEL.visTxtBlkDirection);
        public static Src TextBlockTopMargin => _text_block_cell(CEL.visTxtBlkTopMargin);
        public static Src TextBlockVerticalAlign => _text_block_cell(CEL.visTxtBlkVerticalAlign);

        // Action tags
        private static Src _smart_tag_cell(CEL c) => new Src(SEC.visSectionSmartTag, ROW.visRowSmartTag, c);
        public static Src SmartTagButtonFace => _smart_tag_cell(CEL.visSmartTagButtonFace);
        public static Src SmartTagDescription => _smart_tag_cell(CEL.visSmartTagDescription);
        public static Src SmartTagDisabled => _smart_tag_cell(CEL.visSmartTagDisabled);
        public static Src SmartTagDisplayMode => _smart_tag_cell(CEL.visSmartTagDisplayMode);
        public static Src SmartTagTagName => _smart_tag_cell(CEL.visSmartTagName);
        public static Src SmartTagX => _smart_tag_cell(CEL.visSmartTagX);
        public static Src SmartTagXJustify => _smart_tag_cell(CEL.visSmartTagXJustify);
        public static Src SmartTagY => _smart_tag_cell(CEL.visSmartTagY);
        public static Src SmartTagYJustify => _smart_tag_cell(CEL.visSmartTagYJustify);

        // style
        private static Src _style_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowStyle, c);
        public static Src StyleEnableFillProps => _style_cell(CEL.visStyleIncludesFill);
        public static Src StyleEnableLineProps => _style_cell(CEL.visStyleIncludesLine);
        public static Src StyleEnableTextProps => _style_cell(CEL.visStyleIncludesText);

        //tabs
        private static Src _tab_cell(CEL c) => new Src(SEC.visSectionTab, ROW.visRowTab, c);
        public static Src TabAlignment => _tab_cell(CEL.visTabAlign);
        public static Src TabPosition => _tab_cell(CEL.visTabPos);
        public static Src TabStopCount => _tab_cell(CEL.visTabStopCount);

        // shape layout
        private static Src _shape_layout_cell(CEL c) => new Src(SEC.visSectionObject, ROW.visRowShapeLayout, c);
        public static Src ShapeLayoutConnectorFixedCode => _shape_layout_cell(CEL.visSLOConFixedCode);
        public static Src ShapeLayoutLineJumpCode => _shape_layout_cell(CEL.visSLOJumpCode);
        public static Src ShapeLayoutLineJumpDirX => _shape_layout_cell(CEL.visSLOJumpDirX);
        public static Src ShapeLayoutLineJumpDirY => _shape_layout_cell(CEL.visSLOJumpDirY);
        public static Src ShapeLayoutLineJumpStyle => _shape_layout_cell(CEL.visSLOJumpStyle);
        public static Src ShapeLayoutLineRouteExt => _shape_layout_cell(CEL.visSLOLineRouteExt);
        public static Src ShapeLayoutShapeFixedCode => _shape_layout_cell(CEL.visSLOFixedCode);
        public static Src ShapeLayoutShapePermeablePlace => _shape_layout_cell(CEL.visSLOPermeablePlace);
        public static Src ShapeLayoutShapePermeableX => _shape_layout_cell(CEL.visSLOPermX);
        public static Src ShapeLayoutShapePermeableY => _shape_layout_cell(CEL.visSLOPermY);
        public static Src ShapeLayoutShapePlaceFlip => _shape_layout_cell(CEL.visSLOPlaceFlip);
        public static Src ShapeLayoutShapePlaceStyle => _shape_layout_cell(CEL.visSLOPlaceStyle);
        public static Src ShapeLayoutShapePlowCode => _shape_layout_cell(CEL.visSLOPlowCode);
        public static Src ShapeLayoutShapeRouteStyle => _shape_layout_cell(CEL.visSLORouteStyle);
        public static Src ShapeLayoutShapeSplit => _shape_layout_cell(CEL.visSLOSplit);
        public static Src ShapeLayoutShapeSplittable => _shape_layout_cell(CEL.visSLOSplittable);
        public static Src ShapeLayoutShapeDisplayLevel => _shape_layout_cell(CEL.visSLODisplayLevel); // new in Visio 2010
        public static Src ShapeLayoutRelationships => _shape_layout_cell(CEL.visSLORelationships); // new in Visio 2010
    }
}
