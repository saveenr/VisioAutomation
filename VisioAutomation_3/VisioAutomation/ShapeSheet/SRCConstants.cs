using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using IVisio = Microsoft.Office.Interop.Visio;
using SEC = Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW = Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL = Microsoft.Office.Interop.Visio.VisCellIndices;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static class SRCConstants
    {
        private static SRC.SRCFromCellIndex align = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowAlign);
        private static SRC.SRCFromCellIndex doc = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowDoc);
        private static SRC.SRCFromCellIndex event_ = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowEvent);
        private static SRC.SRCFromCellIndex foreign = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowForeign);
        private static SRC.SRCFromCellIndex fill = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowFill);
        private static SRC.SRCFromCellIndex misc = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowMisc);
        private static SRC.SRCFromCellIndex group_ = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowGroup);
        private static SRC.SRCFromCellIndex image = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowImage);
        private static SRC.SRCFromCellIndex line = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowLine);
        private static SRC.SRCFromCellIndex calendar = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowMisc);
        private static SRC.SRCFromCellIndex oned = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowXForm1D);
        private static SRC.SRCFromCellIndex pagelayout = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowPageLayout);
        private static SRC.SRCFromCellIndex printprops = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowPrintProperties);
        private static SRC.SRCFromCellIndex page = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowPage);
        private static SRC.SRCFromCellIndex lock_ = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowLock);
        private static SRC.SRCFromCellIndex rulergrid = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowRulerGrid);
        private static SRC.SRCFromCellIndex xformout = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowXFormOut);
        private static SRC.SRCFromCellIndex textxfrm = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowTextXForm);
        private static SRC.SRCFromCellIndex text = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowText);
        private static SRC.SRCFromCellIndex style = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowStyle);
        private static SRC.SRCFromCellIndex shapelayout = SRC.GetSRCFactory(SEC.visSectionObject, ROW.visRowShapeLayout);


        private static SRC.SRCFromCellIndex action = SRC.GetSRCFactory(SEC.visSectionAction, ROW.visRowAction);
        private static SRC.SRCFromCellIndex annotation = SRC.GetSRCFactory(SEC.visSectionAnnotation, ROW.visRowAnnotation);
        private static SRC.SRCFromCellIndex char_ = SRC.GetSRCFactory(SEC.visSectionCharacter, ROW.visRowCharacter);
        private static SRC.SRCFromCellIndex connections = SRC.GetSRCFactory(SEC.visSectionConnectionPts, ROW.visRowConnectionPts);
        private static SRC.SRCFromCellIndex controls = SRC.GetSRCFactory(SEC.visSectionControls, ROW.visRowControl);
        private static SRC.SRCFromCellIndex geomvertex = SRC.GetSRCFactory(SEC.visSectionFirstComponent, ROW.visRowVertex);
        private static SRC.SRCFromCellIndex geomcomponent = SRC.GetSRCFactory(SEC.visSectionFirstComponent, ROW.visRowComponent);
        private static SRC.SRCFromCellIndex hyperlink_ = SRC.GetSRCFactory(SEC.visSectionHyperlink, ROW.visRow1stHyperlink);
        private static SRC.SRCFromCellIndex para = SRC.GetSRCFactory(SEC.visSectionParagraph, ROW.visRowParagraph);
        private static SRC.SRCFromCellIndex reviewer = SRC.GetSRCFactory(SEC.visSectionReviewer, ROW.visRowReviewer);
        private static SRC.SRCFromCellIndex prop = SRC.GetSRCFactory(SEC.visSectionProp, ROW.visRowProp);
        private static SRC.SRCFromCellIndex layer = SRC.GetSRCFactory(SEC.visSectionLayer, ROW.visRowLayer);
        private static SRC.SRCFromCellIndex user = SRC.GetSRCFactory(SEC.visSectionUser, ROW.visRowUser);
        private static SRC.SRCFromCellIndex field = SRC.GetSRCFactory(SEC.visSectionTextField, ROW.visRowField);
        private static SRC.SRCFromCellIndex smarttag = SRC.GetSRCFactory(SEC.visSectionSmartTag, ROW.visRowSmartTag);
        private static SRC.SRCFromCellIndex tab = SRC.GetSRCFactory(SEC.visSectionTab, ROW.visRowTab);

        // Actions
        public static readonly SRC Actions_Action = action(CEL.visActionAction);
        public static readonly SRC Actions_BeginGroup = action(CEL.visActionBeginGroup);
        public static readonly SRC Actions_ButtonFace = action(CEL.visActionButtonFace);
        public static readonly SRC Actions_Checked = action(CEL.visActionChecked);
        public static readonly SRC Actions_Disabled = action(CEL.visActionDisabled);
        public static readonly SRC Actions_Invisible = action(CEL.visActionInvisible);
        public static readonly SRC Actions_Menu = action(CEL.visActionMenu);
        public static readonly SRC Actions_ReadOnly = action(CEL.visActionReadOnly);
        public static readonly SRC Actions_SortKey = action(CEL.visActionSortKey);
        public static readonly SRC Actions_TagName = action(CEL.visActionTagName);

        // Alignment
        public static readonly SRC AlignBottom = align(CEL.visAlignBottom);
        public static readonly SRC AlignCenter = align(CEL.visAlignCenter);
        public static readonly SRC AlignLeft = align(CEL.visAlignLeft);
        public static readonly SRC AlignMiddle = align(CEL.visAlignMiddle);
        public static readonly SRC AlignRight = align(CEL.visAlignRight);
        public static readonly SRC AlignTop = align(CEL.visAlignTop);

        // Annotation
        public static readonly SRC Annotation_Comment = annotation(CEL.visAnnotationComment);
        public static readonly SRC Annotation_Date = annotation(CEL.visAnnotationDate);
        public static readonly SRC Annotation_LangID = annotation(CEL.visAnnotationLangID);
        public static readonly SRC Annotation_MarkerIndex = annotation(CEL.visAnnotationMarkerIndex);
        public static readonly SRC Annotation_X = annotation(CEL.visAnnotationX);
        public static readonly SRC Annotation_Y = annotation(CEL.visAnnotationY);

        // Character
        public static readonly SRC Char_AsianFont = char_(CEL.visCharacterAsianFont);
        public static readonly SRC Char_Case = char_(CEL.visCharacterCase);
        public static readonly SRC Char_Color = char_(CEL.visCharacterColor);
        public static readonly SRC Char_ComplexScriptFont = char_(CEL.visCharacterComplexScriptFont);
        public static readonly SRC Char_ComplexScriptSize = char_(CEL.visCharacterComplexScriptSize);
        public static readonly SRC Char_DoubleStrikethrough = char_(CEL.visCharacterDoubleStrikethrough);
        public static readonly SRC Char_DblUnderline = char_(CEL.visCharacterDblUnderline);
        public static readonly SRC Char_Font = char_(CEL.visCharacterFont);
        public static readonly SRC Char_LangID = char_(CEL.visCharacterLangID);
        public static readonly SRC Char_Locale = char_(CEL.visCharacterLocale);
        public static readonly SRC Char_LocalizeFont = char_(CEL.visCharacterLocalizeFont);
        public static readonly SRC Char_Overline = char_(CEL.visCharacterOverline);
        public static readonly SRC Char_Perpendicular = char_(CEL.visCharacterPerpendicular);
        public static readonly SRC Char_Pos = char_(CEL.visCharacterPos);
        public static readonly SRC Char_RTLText = char_(CEL.visCharacterRTLText);
        public static readonly SRC Char_FontScale = char_(CEL.visCharacterFontScale);
        public static readonly SRC Char_Size = char_(CEL.visCharacterSize);
        public static readonly SRC Char_Letterspace = char_(CEL.visCharacterLetterspace);
        public static readonly SRC Char_Strikethru = char_(CEL.visCharacterStrikethru);
        public static readonly SRC Char_Style = char_(CEL.visCharacterStyle);
        public static readonly SRC Char_ColorTrans = char_(CEL.visCharacterColorTrans);

        public static readonly SRC Char_UseVertical = char_(CEL.visCharacterUseVertical);

        // Connections
        public static readonly SRC Connections_D = connections(CEL.visCnnctD);
        public static readonly SRC Connections_DirX = connections(CEL.visCnnctDirX);
        public static readonly SRC Connections_DirY = connections(CEL.visCnnctDirY);
        public static readonly SRC Connections_Type = connections(CEL.visCnnctType);
        public static readonly SRC Connections_X = connections(CEL.visX);
        public static readonly SRC Connections_Y = connections(CEL.visY);

        // Controls
        public static readonly SRC Controls_CanGlue = controls(CEL.visCtlGlue);
        public static readonly SRC Controls_Tip = controls(CEL.visCtlTip);
        public static readonly SRC Controls_XCon = controls(CEL.visCtlXCon);
        public static readonly SRC Controls_X = controls(CEL.visCtlX);
        public static readonly SRC Controls_XDyn = controls(CEL.visCtlXDyn);
        public static readonly SRC Controls_YCon = controls(CEL.visCtlYCon);
        public static readonly SRC Controls_Y = controls(CEL.visCtlY);
        public static readonly SRC Controls_YDyn = controls(CEL.visCtlYDyn);

        // Document Properties

        public static readonly SRC AddMarkup = doc(CEL.visDocAddMarkup);
        public static readonly SRC DocLangID = doc(CEL.visDocLangID);
        public static readonly SRC LockPreview = doc(CEL.visDocLockPreview);
        public static readonly SRC OutputFormat = doc(CEL.visDocOutputFormat);
        public static readonly SRC PreviewQuality = doc(CEL.visDocPreviewQuality);
        public static readonly SRC PreviewScope = doc(CEL.visDocPreviewScope);
        public static readonly SRC ViewMarkup = doc(CEL.visDocViewMarkup);


        // Events
        public static readonly SRC EventDblClick = event_(CEL.visEvtCellDblClick);
        public static readonly SRC EventDrop = event_(CEL.visEvtCellDrop);
        public static readonly SRC EventMultiDrop = event_(CEL.visEvtCellMultiDrop);
        public static readonly SRC EventXFMod = event_(CEL.visEvtCellXFMod);
        public static readonly SRC TheText = event_(CEL.visEvtCellTheText);

        // ForeignImageInfo
        public static readonly SRC ImgHeight = foreign(CEL.visFrgnImgHeight);
        public static readonly SRC ImgOffsetX = foreign(CEL.visFrgnImgOffsetX);
        public static readonly SRC ImgOffsetY = foreign(CEL.visFrgnImgOffsetY);
        public static readonly SRC ImgWidth = foreign(CEL.visFrgnImgWidth);


        // Geometry
        public static readonly SRC Geometry_A = geomvertex(CEL.visBow);
        public static readonly SRC Geometry_B = geomvertex(CEL.visControlX);
        public static readonly SRC Geometry_C = geomvertex(CEL.visEccentricityAngle);
        public static readonly SRC Geometry_D = geomvertex(CEL.visAspectRatio);
        public static readonly SRC Geometry_E = geomvertex(CEL.visNURBSData);
        public static readonly SRC Geometry_X = geomvertex(CEL.visX);
        public static readonly SRC Geometry_Y = geomvertex(CEL.visY);

        public static readonly SRC Geometry_NoFill = geomcomponent(CEL.visCompNoFill);
        public static readonly SRC Geometry_NoLine = geomcomponent(CEL.visCompNoLine);
        public static readonly SRC Geometry_NoShow = geomcomponent(CEL.visCompNoShow);
        public static readonly SRC Geometry_NoSnap = geomcomponent(CEL.visCompNoSnap);


        // Fill Format

        public static readonly SRC FillBkgnd = fill(CEL.visFillBkgnd);
        public static readonly SRC FillBkgndTrans = fill(CEL.visFillBkgndTrans);
        public static readonly SRC FillForegnd = fill(CEL.visFillForegnd);
        public static readonly SRC FillForegndTrans = fill(CEL.visFillForegndTrans);
        public static readonly SRC FillPattern = fill(CEL.visFillPattern);
        public static readonly SRC ShapeShdwObliqueAngle = fill(CEL.visFillShdwObliqueAngle);
        public static readonly SRC ShapeShdwOffsetX = fill(CEL.visFillShdwOffsetX);
        public static readonly SRC ShapeShdwOffsetY = fill(CEL.visFillShdwOffsetY);
        public static readonly SRC ShapeShdwScaleFactor = fill(CEL.visFillShdwScaleFactor);
        public static readonly SRC ShapeShdwType = fill(CEL.visFillShdwType);
        public static readonly SRC ShdwBkgnd = fill(CEL.visFillShdwBkgnd);
        public static readonly SRC ShdwBkgndTrans = fill(CEL.visFillShdwBkgndTrans);
        public static readonly SRC ShdwForegnd = fill(CEL.visFillShdwForegnd);
        public static readonly SRC ShdwForegndTrans = fill(CEL.visFillShdwForegndTrans);
        public static readonly SRC ShdwPattern = fill(CEL.visFillShdwPattern);

        // GlueInfo
        public static readonly SRC BegTrigger = misc(CEL.visBegTrigger);
        public static readonly SRC EndTrigger = misc(CEL.visEndTrigger);
        public static readonly SRC GlueType = misc(CEL.visGlueType);
        public static readonly SRC WalkPreference = misc(CEL.visWalkPref);

        // GroupProperties

        public static readonly SRC DisplayMode = group_(CEL.visGroupDisplayMode);
        public static readonly SRC DontMoveChildren = group_(CEL.visGroupDontMoveChildren);
        public static readonly SRC IsDropTarget = group_(CEL.visGroupIsDropTarget);
        public static readonly SRC IsSnapTarget = group_(CEL.visGroupIsSnapTarget);
        public static readonly SRC IsTextEditTarget = group_(CEL.visGroupIsTextEditTarget);
        public static readonly SRC SelectMode = group_(CEL.visGroupSelectMode);

        // Hyperlinks

        public static readonly SRC Hyperlink_Address = hyperlink_(CEL.visHLinkAddress);
        public static readonly SRC Hyperlink_Default = hyperlink_(CEL.visHLinkDefault);
        public static readonly SRC Hyperlink_Description = hyperlink_(CEL.visHLinkDescription);
        public static readonly SRC Hyperlink_ExtraInfo = hyperlink_(CEL.visHLinkExtraInfo);
        public static readonly SRC Hyperlink_Frame = hyperlink_(CEL.visHLinkFrame);
        public static readonly SRC Hyperlink_Invisible = hyperlink_(CEL.visHLinkInvisible);
        public static readonly SRC Hyperlink_NewWindow = hyperlink_(CEL.visHLinkNewWin);
        public static readonly SRC Hyperlink_SortKey = hyperlink_(CEL.visHLinkSortKey);
        public static readonly SRC Hyperlink_SubAddress = hyperlink_(CEL.visHLinkSubAddress);

        // Image Properties
        public static readonly SRC Blur = image(CEL.visImageBlur);
        public static readonly SRC Brightness = image(CEL.visImageBrightness);
        public static readonly SRC Contrast = image(CEL.visImageContrast);
        public static readonly SRC Denoise = image(CEL.visImageDenoise);
        public static readonly SRC Gamma = image(CEL.visImageGamma);
        public static readonly SRC Sharpen = image(CEL.visImageSharpen);
        public static readonly SRC Transparency = image(CEL.visImageTransparency);


        // Line format
        public static readonly SRC BeginArrow = line(CEL.visLineBeginArrow);
        public static readonly SRC BeginArrowSize = line(CEL.visLineBeginArrowSize);
        public static readonly SRC EndArrow = line(CEL.visLineEndArrow);
        public static readonly SRC EndArrowSize = line(CEL.visLineEndArrowSize);
        public static readonly SRC LineCap = line(CEL.visLineEndCap);
        public static readonly SRC LineColor = line(CEL.visLineColor);
        public static readonly SRC LineColorTrans = line(CEL.visLineColorTrans);
        public static readonly SRC LinePattern = line(CEL.visLinePattern);
        public static readonly SRC LineWeight = line(CEL.visLineWeight);
        public static readonly SRC Rounding = line(CEL.visLineRounding);


        // Miscellaneous

        public static readonly SRC Calendar = calendar(CEL.visObjCalendar);
        public static readonly SRC Comment = calendar(CEL.visComment);
        public static readonly SRC DropOnPageScale = calendar(CEL.visObjDropOnPageScale);
        public static readonly SRC DynFeedback = calendar(CEL.visDynFeedback);
        public static readonly SRC IsDropSource = calendar(CEL.visDropSource);
        public static readonly SRC LangID = calendar(CEL.visObjLangID);
        public static readonly SRC LocalizeMerge = calendar(CEL.visObjLocalizeMerge);
        public static readonly SRC NoAlignBox = calendar(CEL.visNoAlignBox);
        public static readonly SRC NoCtlHandles = calendar(CEL.visNoCtlHandles);
        public static readonly SRC NoLiveDynamics = calendar(CEL.visNoLiveDynamics);
        public static readonly SRC NonPrinting = calendar(CEL.visNonPrinting);
        public static readonly SRC NoObjHandles = calendar(CEL.visNoObjHandles);
        public static readonly SRC ObjType = calendar(CEL.visLOFlags);
        public static readonly SRC UpdateAlignBox = calendar(CEL.visUpdateAlignBox);

        // 1d endpoints

        public static readonly SRC BeginX = oned(CEL.vis1DBeginX);
        public static readonly SRC BeginY = oned(CEL.vis1DBeginY);
        public static readonly SRC EndX = oned(CEL.vis1DEndX);
        public static readonly SRC EndY = oned(CEL.vis1DEndY);


        // page layout
        public static readonly SRC AvenueSizeX = pagelayout(CEL.visPLOAvenueSizeX);
        public static readonly SRC AvenueSizeY = pagelayout(CEL.visPLOAvenueSizeY);
        public static readonly SRC BlockSizeX = pagelayout(CEL.visPLOBlockSizeX);
        public static readonly SRC BlockSizeY = pagelayout(CEL.visPLOBlockSizeY);
        public static readonly SRC CtrlAsInput = pagelayout(CEL.visPLOCtrlAsInput);
        public static readonly SRC DynamicsOff = pagelayout(CEL.visPLODynamicsOff);
        public static readonly SRC EnableGrid = pagelayout(CEL.visPLOEnableGrid);
        public static readonly SRC LineAdjustFrom = pagelayout(CEL.visPLOLineAdjustFrom);
        public static readonly SRC LineAdjustTo = pagelayout(CEL.visPLOLineAdjustTo);
        public static readonly SRC LineJumpCode = pagelayout(CEL.visPLOJumpCode);
        public static readonly SRC LineJumpFactorX = pagelayout(CEL.visPLOJumpFactorX);
        public static readonly SRC LineJumpFactorY = pagelayout(CEL.visPLOJumpFactorY);
        public static readonly SRC LineJumpStyle = pagelayout(CEL.visPLOJumpStyle);
        public static readonly SRC LineRouteExt = pagelayout(CEL.visPLOLineRouteExt);
        public static readonly SRC LineToLineX = pagelayout(CEL.visPLOLineToLineX);
        public static readonly SRC LineToLineY = pagelayout(CEL.visPLOLineToLineY);
        public static readonly SRC LineToNodeX = pagelayout(CEL.visPLOLineToNodeX);
        public static readonly SRC LineToNodeY = pagelayout(CEL.visPLOLineToNodeY);
        public static readonly SRC PageLineJumpDirX = pagelayout(CEL.visPLOJumpDirX);
        public static readonly SRC PageLineJumpDirY = pagelayout(CEL.visPLOJumpDirY);
        public static readonly SRC PageShapeSplit = pagelayout(CEL.visPLOSplit);
        public static readonly SRC PlaceDepth = pagelayout(CEL.visPLOPlaceDepth);
        public static readonly SRC PlaceFlip = pagelayout(CEL.visPLOPlaceFlip);
        public static readonly SRC PlaceStyle = pagelayout(CEL.visPLOPlaceStyle);
        public static readonly SRC PlowCode = pagelayout(CEL.visPLOPlowCode);
        public static readonly SRC ResizePage = pagelayout(CEL.visPLOResizePage);
        public static readonly SRC RouteStyle = pagelayout(CEL.visPLORouteStyle);


        // print properties

        public static readonly SRC PageLeftMargin = printprops(CEL.visPrintPropertiesLeftMargin);
        public static readonly SRC CenterX = printprops(CEL.visPrintPropertiesCenterX);
        public static readonly SRC CenterY = printprops(CEL.visPrintPropertiesCenterY);
        public static readonly SRC OnPage = printprops(CEL.visPrintPropertiesOnPage);
        public static readonly SRC PageBottomMargin = printprops(CEL.visPrintPropertiesBottomMargin);
        public static readonly SRC PageRightMargin = printprops(CEL.visPrintPropertiesRightMargin);
        public static readonly SRC PagesX = printprops(CEL.visPrintPropertiesPagesX);
        public static readonly SRC PagesY = printprops(CEL.visPrintPropertiesPagesY);
        public static readonly SRC PageTopMargin = printprops(CEL.visPrintPropertiesTopMargin);
        public static readonly SRC PaperKind = printprops(CEL.visPrintPropertiesPaperKind);
        public static readonly SRC PrintGrid = printprops(CEL.visPrintPropertiesPrintGrid);
        public static readonly SRC PrintPageOrientation = printprops(CEL.visPrintPropertiesPageOrientation);
        public static readonly SRC ScaleX = printprops(CEL.visPrintPropertiesScaleX);
        public static readonly SRC ScaleY = printprops(CEL.visPrintPropertiesScaleY);
        public static readonly SRC PaperSource = printprops(CEL.visPrintPropertiesPaperSource);

        // page properties

        public static readonly SRC DrawingScale = page(CEL.visPageDrawingScale);
        public static readonly SRC DrawingScaleType = page(CEL.visPageDrawScaleType);
        public static readonly SRC DrawingSizeType = page(CEL.visPageDrawSizeType);
        public static readonly SRC InhibitSnap = page(CEL.visPageInhibitSnap);
        public static readonly SRC PageHeight = page(CEL.visPageHeight);
        public static readonly SRC PageScale = page(CEL.visPageScale);
        public static readonly SRC PageWidth = page(CEL.visPageWidth);
        public static readonly SRC ShdwObliqueAngle = page(CEL.visPageShdwObliqueAngle);
        public static readonly SRC ShdwOffsetX = page(CEL.visPageShdwOffsetX);
        public static readonly SRC ShdwOffsetY = page(CEL.visPageShdwOffsetY);
        public static readonly SRC ShdwScaleFactor = page(CEL.visPageShdwScaleFactor);
        public static readonly SRC ShdwType = page(CEL.visPageShdwType);
        public static readonly SRC UIVisibility = page(CEL.visPageUIVisibility);

        // paragraph
        public static readonly SRC Para_Bullet = para(CEL.visBulletIndex);
        public static readonly SRC Para_BulletFont = para(CEL.visBulletFont);
        public static readonly SRC Para_BulletFontSize = para(CEL.visBulletFontSize);
        public static readonly SRC Para_BulletStr = para(CEL.visBulletString);
        public static readonly SRC Para_Flags = para(CEL.visFlags);
        public static readonly SRC Para_HorzAlign = para(CEL.visHorzAlign);
        public static readonly SRC Para_IndFirst = para(CEL.visIndentFirst);
        public static readonly SRC Para_IndLeft = para(CEL.visIndentLeft);
        public static readonly SRC Para_IndRight = para(CEL.visIndentRight);
        public static readonly SRC Para_LocalizeBulletFont = para(CEL.visLocalizeBulletFont);
        public static readonly SRC Para_SpAfter = para(CEL.visSpaceAfter);
        public static readonly SRC Para_SpBefore = para(CEL.visSpaceBefore);
        public static readonly SRC Para_SpLine = para(CEL.visSpaceLine);
        public static readonly SRC Para_TextPosAfterBullet = para(CEL.visTextPosAfterBullet);

        // protection

        public static readonly SRC LockAspect = lock_(CEL.visLockAspect);
        public static readonly SRC LockBegin = lock_(CEL.visLockBegin);
        public static readonly SRC LockCalcWH = lock_(CEL.visLockCalcWH);
        public static readonly SRC LockCrop = lock_(CEL.visLockCrop);
        public static readonly SRC LockCustProp = lock_(CEL.visLockCustProp);
        public static readonly SRC LockDelete = lock_(CEL.visLockDelete);
        public static readonly SRC LockEnd = lock_(CEL.visLockEnd);
        public static readonly SRC LockFormat = lock_(CEL.visLockFormat);
        public static readonly SRC LockFromGroupFormat = lock_(CEL.visLockFromGroupFormat);
        public static readonly SRC LockGroup = lock_(CEL.visLockGroup);
        public static readonly SRC LockHeight = lock_(CEL.visLockHeight);
        public static readonly SRC LockMoveX = lock_(CEL.visLockMoveX);
        public static readonly SRC LockMoveY = lock_(CEL.visLockMoveY);
        public static readonly SRC LockRotate = lock_(CEL.visLockRotate);
        public static readonly SRC LockSelect = lock_(CEL.visLockSelect);
        public static readonly SRC LockTextEdit = lock_(CEL.visLockTextEdit);
        public static readonly SRC LockThemeColors = lock_(CEL.visLockThemeColors);
        public static readonly SRC LockThemeEffects = lock_(CEL.visLockThemeEffects);
        public static readonly SRC LockVtxEdit = lock_(CEL.visLockVtxEdit);
        public static readonly SRC LockWidth = lock_(CEL.visLockWidth);


        // ruler and grid

        public static readonly SRC XGridDensity = rulergrid(CEL.visXGridDensity);
        public static readonly SRC XGridOrigin = rulergrid(CEL.visXGridOrigin);
        public static readonly SRC XGridSpacing = rulergrid(CEL.visXGridSpacing);
        public static readonly SRC XRulerDensity = rulergrid(CEL.visXRulerDensity);
        public static readonly SRC XRulerOrigin = rulergrid(CEL.visXRulerOrigin);
        public static readonly SRC YGridDensity = rulergrid(CEL.visYGridDensity);
        public static readonly SRC YGridOrigin = rulergrid(CEL.visYGridOrigin);
        public static readonly SRC YGridSpacing = rulergrid(CEL.visYGridSpacing);
        public static readonly SRC YRulerDensity = rulergrid(CEL.visYRulerDensity);
        public static readonly SRC YRulerOrigin = rulergrid(CEL.visYRulerOrigin);


        // Shape Tranform

        public static readonly SRC Angle = xformout(CEL.visXFormAngle);
        public static readonly SRC FlipX = xformout(CEL.visXFormFlipX);
        public static readonly SRC FlipY = xformout(CEL.visXFormFlipY);
        public static readonly SRC Height = xformout(CEL.visXFormHeight);
        public static readonly SRC LocPinX = xformout(CEL.visXFormLocPinX);
        public static readonly SRC LocPinY = xformout(CEL.visXFormLocPinY);
        public static readonly SRC PinX = xformout(CEL.visXFormPinX);
        public static readonly SRC PinY = xformout(CEL.visXFormPinY);
        public static readonly SRC ResizeMode = xformout(CEL.visXFormResizeMode);
        public static readonly SRC Width = xformout(CEL.visXFormWidth);

        // reviewer

        public static readonly SRC Reviewer_Color = reviewer(CEL.visReviewerColor);
        public static readonly SRC Reviewer_Initials = reviewer(CEL.visReviewerInitials);
        public static readonly SRC Reviewer_Name = reviewer(CEL.visReviewerName);

        // shape data

        public static readonly SRC Prop_SortKey = prop(CEL.visCustPropsSortKey);
        public static readonly SRC Prop_Ask = prop(CEL.visCustPropsAsk);
        public static readonly SRC Prop_Calendar = prop(CEL.visCustPropsCalendar);
        public static readonly SRC Prop_Format = prop(CEL.visCustPropsFormat);
        public static readonly SRC Prop_Invisible = prop(CEL.visCustPropsInvis);
        public static readonly SRC Prop_Label = prop(CEL.visCustPropsLabel);
        public static readonly SRC Prop_LangID = prop(CEL.visCustPropsLangID);
        public static readonly SRC Prop_Prompt = prop(CEL.visCustPropsPrompt);
        public static readonly SRC Prop_Type = prop(CEL.visCustPropsType);
        public static readonly SRC Prop_Value = prop(CEL.visCustPropsValue);

        // Layers

        public static readonly SRC Layers_Active = layer(CEL.visLayerActive);
        public static readonly SRC Layers_Color = layer(CEL.visLayerColor);
        public static readonly SRC Layers_Glue = layer(CEL.visLayerGlue);
        public static readonly SRC Layers_Locked = layer(CEL.visLayerLock);
        public static readonly SRC Layers_Print = layer(CEL.visDocPreviewScope);
        public static readonly SRC Layers_Snap = layer(CEL.visLayerSnap);
        public static readonly SRC Layers_ColorTrans = layer(CEL.visLayerColorTrans);
        public static readonly SRC Layers_Visible = layer(CEL.visLayerVisible);


        //text transform
        public static readonly SRC TxtAngle = textxfrm(CEL.visXFormAngle);
        public static readonly SRC TxtHeight = textxfrm(CEL.visXFormHeight);
        public static readonly SRC TxtLocPinX = textxfrm(CEL.visXFormLocPinX);
        public static readonly SRC TxtLocPinY = textxfrm(CEL.visXFormLocPinY);
        public static readonly SRC TxtPinX = textxfrm(CEL.visXFormPinX);
        public static readonly SRC TxtPinY = textxfrm(CEL.visXFormPinY);
        public static readonly SRC TxtWidth = textxfrm(CEL.visXFormWidth);


        // user defined cells
        public static readonly SRC User_Prompt = user(CEL.visUserPrompt);
        public static readonly SRC User_Value = user(CEL.visUserValue);


        // Fields
        public static readonly SRC Fields_Calendar = field(CEL.visFieldCalendar);
        public static readonly SRC Fields_Format = field(CEL.visFieldFormat);
        public static readonly SRC Fields_ObjectKind = field(CEL.visFieldObjectKind);
        public static readonly SRC Fields_Type = field(CEL.visFieldType);
        public static readonly SRC Fields_UICat = field(CEL.visFieldUICategory);
        public static readonly SRC Fields_UICod = field(CEL.visFieldUICode);
        public static readonly SRC Fields_UIFmt = field(CEL.visFieldUIFormat);
        public static readonly SRC Fields_Value = field(CEL.visFieldCell);


        // text block format

        public static readonly SRC BottomMargin = text(CEL.visTxtBlkBottomMargin);
        public static readonly SRC DefaultTabStop = text(CEL.visTxtBlkDefaultTabStop);
        public static readonly SRC LeftMargin = text(CEL.visTxtBlkLeftMargin);
        public static readonly SRC RightMargin = text(CEL.visTxtBlkRightMargin);
        public static readonly SRC TextBkgnd = text(CEL.visTxtBlkBkgnd);
        public static readonly SRC TextBkgndTrans = text(CEL.visTxtBlkBkgndTrans);
        public static readonly SRC TextDirection = text(CEL.visTxtBlkDirection);
        public static readonly SRC TopMargin = text(CEL.visTxtBlkTopMargin);
        public static readonly SRC VerticalAlign = text(CEL.visTxtBlkVerticalAlign);

        // Action tags
        public static readonly SRC SmartTags_ButtonFace = smarttag(CEL.visSmartTagButtonFace);
        public static readonly SRC SmartTags_Description = smarttag(CEL.visSmartTagDescription);
        public static readonly SRC SmartTags_Disabled = smarttag(CEL.visSmartTagDisabled);
        public static readonly SRC SmartTags_DisplayMode = smarttag(CEL.visSmartTagDisplayMode);
        public static readonly SRC SmartTags_TagName = smarttag(CEL.visSmartTagName);
        public static readonly SRC SmartTags_X = smarttag(CEL.visSmartTagX);
        public static readonly SRC SmartTags_XJustify = smarttag(CEL.visSmartTagXJustify);
        public static readonly SRC SmartTags_Y = smarttag(CEL.visSmartTagY);
        public static readonly SRC SmartTags_YJustify = smarttag(CEL.visSmartTagYJustify);

        // style
        public static readonly SRC EnableFillProps = style(CEL.visStyleIncludesFill);
        public static readonly SRC EnableLineProps = style(CEL.visStyleIncludesLine);
        public static readonly SRC EnableTextProps = style(CEL.visStyleIncludesText);
        public static readonly SRC HideText = style(CEL.visStyleHidden);


        //tabs
        public static readonly SRC Tabs_Alignment = tab(CEL.visTabAlign);
        public static readonly SRC Tabs_Position = tab(CEL.visTabPos);
        public static readonly SRC Tabs_StopCount = tab(CEL.visTabStopCount);

        // shape layout
        public static readonly SRC ConFixedCode = shapelayout(CEL.visSLOConFixedCode);
        public static readonly SRC ConLineJumpCode = shapelayout(CEL.visSLOJumpCode);
        public static readonly SRC ConLineJumpDirX = shapelayout(CEL.visSLOJumpDirX);
        public static readonly SRC ConLineJumpDirY = shapelayout(CEL.visSLOJumpDirY);
        public static readonly SRC ConLineJumpStyle = shapelayout(CEL.visSLOJumpStyle);
        public static readonly SRC ConLineRouteExt = shapelayout(CEL.visSLOLineRouteExt);
        public static readonly SRC ShapeFixedCode = shapelayout(CEL.visSLOFixedCode);
        public static readonly SRC ShapePermeablePlace = shapelayout(CEL.visSLOPermeablePlace);
        public static readonly SRC ShapePermeableX = shapelayout(CEL.visSLOPermX);
        public static readonly SRC ShapePermeableY = shapelayout(CEL.visSLOPermY);
        public static readonly SRC ShapePlaceFlip = shapelayout(CEL.visSLOPlaceFlip);
        public static readonly SRC ShapePlaceStyle = shapelayout(CEL.visSLOPlaceStyle);
        public static readonly SRC ShapePlowCode = shapelayout(CEL.visSLOPlowCode);
        public static readonly SRC ShapeRouteStyle = shapelayout(CEL.visSLORouteStyle);
        public static readonly SRC ShapeSplit = shapelayout(CEL.visSLOSplit);
        public static readonly SRC ShapeSplittable = shapelayout(CEL.visSLOSplittable);

        // Static Methods

        public static Dictionary<string, VA.ShapeSheet.SRC> GetSRCDictionary()
        {
            var fields = GetSRCFields();

            var fields_name_to_value = new Dictionary<string, VA.ShapeSheet.SRC>();
            foreach (var field in fields)
            {
                fields_name_to_value[field.Name] = (VA.ShapeSheet.SRC) field.GetValue(null);
            }

            return fields_name_to_value;
        }

        private static List<FieldInfo> GetSRCFields()
        {
            var srcconstants_t = typeof (VA.ShapeSheet.SRCConstants);
            var fields = srcconstants_t.GetFields()
                .Where(m => m.FieldType == typeof (VA.ShapeSheet.SRC))
                .Where(m => m.IsPublic)
                .Where(m => m.IsStatic)
                .ToList();
            return fields;
        }
    }
}