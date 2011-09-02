using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using IVisio = Microsoft.Office.Interop.Visio;
using SEC=Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW= Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL= Microsoft.Office.Interop.Visio.VisCellIndices;
using VA=VisioAutomation;

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
        public readonly static SRC Actions_Action = action(CEL.visActionAction);
        public readonly static SRC Actions_BeginGroup = action(CEL.visActionBeginGroup);
        public readonly static SRC Actions_ButtonFace = action(CEL.visActionButtonFace);
        public readonly static SRC Actions_Checked = action(CEL.visActionChecked);
        public readonly static SRC Actions_Disabled = action(CEL.visActionDisabled);
        public readonly static SRC Actions_Invisible = action(CEL.visActionInvisible);
        public readonly static SRC Actions_Menu = action(CEL.visActionMenu);
        public readonly static SRC Actions_ReadOnly = action(CEL.visActionReadOnly);
        public readonly static SRC Actions_SortKey = action(CEL.visActionSortKey);
        public readonly static SRC Actions_TagName = action(CEL.visActionTagName);

        // Alignment
        public readonly static SRC AlignBottom = align( CEL.visAlignBottom);
        public readonly static SRC AlignCenter = align( CEL.visAlignCenter);
        public readonly static SRC AlignLeft = align( CEL.visAlignLeft);
        public readonly static SRC AlignMiddle = align( CEL.visAlignMiddle);
        public readonly static SRC AlignRight = align( CEL.visAlignRight);
        public readonly static SRC AlignTop = align( CEL.visAlignTop);

        // Annotation
        public readonly static SRC Annotation_Comment = annotation(CEL.visAnnotationComment);
        public readonly static SRC Annotation_Date = annotation(CEL.visAnnotationDate);
        public readonly static SRC Annotation_LangID = annotation(CEL.visAnnotationLangID);
        public readonly static SRC Annotation_MarkerIndex = annotation(CEL.visAnnotationMarkerIndex);
        public readonly static SRC Annotation_X = annotation(CEL.visAnnotationX);
        public readonly static SRC Annotation_Y = annotation(CEL.visAnnotationY);

        // Character
        public readonly static SRC Char_AsianFont = char_( CEL.visCharacterAsianFont);
        public readonly static SRC Char_Case = char_( CEL.visCharacterCase);
        public readonly static SRC Char_Color = char_( CEL.visCharacterColor);
        public readonly static SRC Char_ComplexScriptFont = char_( CEL.visCharacterComplexScriptFont);
        public readonly static SRC Char_ComplexScriptSize = char_( CEL.visCharacterComplexScriptSize);
        public readonly static SRC Char_DoubleStrikethrough = char_( CEL.visCharacterDoubleStrikethrough);
        public readonly static SRC Char_DblUnderline = char_( CEL.visCharacterDblUnderline);
        public readonly static SRC Char_Font = char_( CEL.visCharacterFont);
        public readonly static SRC Char_LangID = char_( CEL.visCharacterLangID);
        public readonly static SRC Char_Locale = char_( CEL.visCharacterLocale);
        public readonly static SRC Char_LocalizeFont = char_( CEL.visCharacterLocalizeFont);
        public readonly static SRC Char_Overline = char_( CEL.visCharacterOverline);
        public readonly static SRC Char_Perpendicular = char_( CEL.visCharacterPerpendicular);
        public readonly static SRC Char_Pos = char_( CEL.visCharacterPos);
        public readonly static SRC Char_RTLText = char_( CEL.visCharacterRTLText);
        public readonly static SRC Char_FontScale = char_( CEL.visCharacterFontScale);
        public readonly static SRC Char_Size = char_( CEL.visCharacterSize);
        public readonly static SRC Char_Letterspace = char_( CEL.visCharacterLetterspace);
        public readonly static SRC Char_Strikethru = char_( CEL.visCharacterStrikethru);
        public readonly static SRC Char_Style = char_( CEL.visCharacterStyle);
        public readonly static SRC Char_ColorTrans = char_( CEL.visCharacterColorTrans);
        
        public readonly static SRC Char_UseVertical = char_( CEL.visCharacterUseVertical);

        // Connections
        public readonly static SRC Connections_D = connections(CEL.visCnnctD);
        public readonly static SRC Connections_DirX = connections(CEL.visCnnctDirX);
        public readonly static SRC Connections_DirY = connections(CEL.visCnnctDirY);
        public readonly static SRC Connections_Type = connections(CEL.visCnnctType);
        public readonly static SRC Connections_X = connections(CEL.visX);
        public readonly static SRC Connections_Y = connections(CEL.visY);

        // Controls
        public readonly static SRC Controls_CanGlue = controls( CEL.visCtlGlue);
        public readonly static SRC Controls_Tip = controls( CEL.visCtlTip);
        public readonly static SRC Controls_XCon = controls( CEL.visCtlXCon);
        public readonly static SRC Controls_X = controls( CEL.visCtlX);
        public readonly static SRC Controls_XDyn = controls( CEL.visCtlXDyn);
        public readonly static SRC Controls_YCon = controls( CEL.visCtlYCon);
        public readonly static SRC Controls_Y = controls( CEL.visCtlY);
        public readonly static SRC Controls_YDyn = controls( CEL.visCtlYDyn);

        // Document Properties

        public readonly static SRC AddMarkup = doc(CEL.visDocAddMarkup);
        public readonly static SRC DocLangID = doc(CEL.visDocLangID);
        public readonly static SRC LockPreview = doc(CEL.visDocLockPreview);
        public readonly static SRC OutputFormat = doc(CEL.visDocOutputFormat);
        public readonly static SRC PreviewQuality = doc(CEL.visDocPreviewQuality);
        public readonly static SRC PreviewScope = doc(CEL.visDocPreviewScope);
        public readonly static SRC ViewMarkup = doc(CEL.visDocViewMarkup);


        // Events
        public readonly static SRC EventDblClick = event_(CEL.visEvtCellDblClick);
        public readonly static SRC EventDrop = event_(CEL.visEvtCellDrop);
        public readonly static SRC EventMultiDrop = event_(CEL.visEvtCellMultiDrop);
        public readonly static SRC EventXFMod = event_(CEL.visEvtCellXFMod);
        public readonly static SRC TheText = event_(CEL.visEvtCellTheText);

        // ForeignImageInfo
        public readonly static SRC ImgHeight =  foreign( CEL.visFrgnImgHeight);
        public readonly static SRC ImgOffsetX = foreign( CEL.visFrgnImgOffsetX);
        public readonly static SRC ImgOffsetY = foreign( CEL.visFrgnImgOffsetY);
        public readonly static SRC ImgWidth =   foreign( CEL.visFrgnImgWidth);


        // Geometry
        public readonly static SRC Geometry_A =      geomvertex( CEL.visBow);
        public readonly static SRC Geometry_B =      geomvertex( CEL.visControlX);
        public readonly static SRC Geometry_C =      geomvertex( CEL.visEccentricityAngle);
        public readonly static SRC Geometry_D =      geomvertex( CEL.visAspectRatio);
        public readonly static SRC Geometry_E =      geomvertex( CEL.visNURBSData);
        public readonly static SRC Geometry_X =      geomvertex( CEL.visX);
        public readonly static SRC Geometry_Y =      geomvertex( CEL.visY);

        public readonly static SRC Geometry_NoFill = geomcomponent(  CEL.visCompNoFill);
        public readonly static SRC Geometry_NoLine = geomcomponent(  CEL.visCompNoLine);
        public readonly static SRC Geometry_NoShow = geomcomponent(  CEL.visCompNoShow);
        public readonly static SRC Geometry_NoSnap = geomcomponent(  CEL.visCompNoSnap);
 

        // Fill Format

        public readonly static SRC FillBkgnd = fill(CEL.visFillBkgnd);
        public readonly static SRC FillBkgndTrans = fill(CEL.visFillBkgndTrans);
        public readonly static SRC FillForegnd = fill(CEL.visFillForegnd);
        public readonly static SRC FillForegndTrans = fill(CEL.visFillForegndTrans);
        public readonly static SRC FillPattern = fill(CEL.visFillPattern);
        public readonly static SRC ShapeShdwObliqueAngle = fill(CEL.visFillShdwObliqueAngle);
        public readonly static SRC ShapeShdwOffsetX = fill(CEL.visFillShdwOffsetX);
        public readonly static SRC ShapeShdwOffsetY = fill(CEL.visFillShdwOffsetY);
        public readonly static SRC ShapeShdwScaleFactor = fill(CEL.visFillShdwScaleFactor);
        public readonly static SRC ShapeShdwType = fill(CEL.visFillShdwType);
        public readonly static SRC ShdwBkgnd = fill(CEL.visFillShdwBkgnd);
        public readonly static SRC ShdwBkgndTrans = fill(CEL.visFillShdwBkgndTrans);
        public readonly static SRC ShdwForegnd = fill(CEL.visFillShdwForegnd);
        public readonly static SRC ShdwForegndTrans = fill(CEL.visFillShdwForegndTrans);
        public readonly static SRC ShdwPattern = fill(CEL.visFillShdwPattern);

        // GlueInfo
        public readonly static SRC BegTrigger =     misc( CEL.visBegTrigger);
        public readonly static SRC EndTrigger =     misc( CEL.visEndTrigger);
        public readonly static SRC GlueType =       misc( CEL.visGlueType);
        public readonly static SRC WalkPreference = misc( CEL.visWalkPref);

        // GroupProperties

        public readonly static SRC DisplayMode =      group_( CEL.visGroupDisplayMode);
        public readonly static SRC DontMoveChildren = group_( CEL.visGroupDontMoveChildren);
        public readonly static SRC IsDropTarget =     group_( CEL.visGroupIsDropTarget);
        public readonly static SRC IsSnapTarget =     group_( CEL.visGroupIsSnapTarget);
        public readonly static SRC IsTextEditTarget = group_( CEL.visGroupIsTextEditTarget);
        public readonly static SRC SelectMode =       group_( CEL.visGroupSelectMode);

        // Hyperlinks

        public readonly static SRC Hyperlink_Address =    hyperlink_( CEL.visHLinkAddress);
        public readonly static SRC Hyperlink_Default =    hyperlink_( CEL.visHLinkDefault);
        public readonly static SRC Hyperlink_Description =hyperlink_( CEL.visHLinkDescription);
        public readonly static SRC Hyperlink_ExtraInfo =  hyperlink_( CEL.visHLinkExtraInfo);
        public readonly static SRC Hyperlink_Frame =      hyperlink_( CEL.visHLinkFrame);
        public readonly static SRC Hyperlink_Invisible =  hyperlink_( CEL.visHLinkInvisible);
        public readonly static SRC Hyperlink_NewWindow =  hyperlink_( CEL.visHLinkNewWin);
        public readonly static SRC Hyperlink_SortKey =    hyperlink_( CEL.visHLinkSortKey);
        public readonly static SRC Hyperlink_SubAddress = hyperlink_( CEL.visHLinkSubAddress);

        // Image Properties
        public readonly static SRC Blur = image( CEL.visImageBlur);
        public readonly static SRC Brightness = image( CEL.visImageBrightness);
        public readonly static SRC Contrast = image( CEL.visImageContrast);
        public readonly static SRC Denoise = image( CEL.visImageDenoise);
        public readonly static SRC Gamma = image( CEL.visImageGamma);
        public readonly static SRC Sharpen = image( CEL.visImageSharpen);
        public readonly static SRC Transparency = image( CEL.visImageTransparency);


        // Line format
        public readonly static SRC BeginArrow = line(CEL.visLineBeginArrow);
        public readonly static SRC BeginArrowSize = line(CEL.visLineBeginArrowSize);
        public readonly static SRC EndArrow = line(CEL.visLineEndArrow);
        public readonly static SRC EndArrowSize = line(CEL.visLineEndArrowSize);
        public readonly static SRC LineCap = line(CEL.visLineEndCap);
        public readonly static SRC LineColor = line(CEL.visLineColor);
        public readonly static SRC LineColorTrans = line(CEL.visLineColorTrans);
        public readonly static SRC LinePattern = line(CEL.visLinePattern);
        public readonly static SRC LineWeight = line(CEL.visLineWeight);
        public readonly static SRC Rounding = line(CEL.visLineRounding);


        // Miscellaneous

        public readonly static SRC Calendar =calendar( CEL.visObjCalendar);
        public readonly static SRC Comment =calendar( CEL.visComment);
        public readonly static SRC DropOnPageScale =calendar( CEL.visObjDropOnPageScale);
        public readonly static SRC DynFeedback =calendar( CEL.visDynFeedback);
        public readonly static SRC IsDropSource =calendar( CEL.visDropSource);
        public readonly static SRC LangID =calendar( CEL.visObjLangID);
        public readonly static SRC LocalizeMerge =calendar( CEL.visObjLocalizeMerge);
        public readonly static SRC NoAlignBox =calendar( CEL.visNoAlignBox);
        public readonly static SRC NoCtlHandles =calendar( CEL.visNoCtlHandles);
        public readonly static SRC NoLiveDynamics =calendar( CEL.visNoLiveDynamics);
        public readonly static SRC NonPrinting =calendar( CEL.visNonPrinting);
        public readonly static SRC NoObjHandles =calendar( CEL.visNoObjHandles);
        public readonly static SRC ObjType =calendar( CEL.visLOFlags);
        public readonly static SRC UpdateAlignBox =calendar( CEL.visUpdateAlignBox);

        // 1d endpoints

        public readonly static SRC BeginX = oned( CEL.vis1DBeginX);
        public readonly static SRC BeginY = oned( CEL.vis1DBeginY);
        public readonly static SRC EndX =   oned( CEL.vis1DEndX);
        public readonly static SRC EndY =   oned( CEL.vis1DEndY);


        // page layout
        public readonly static SRC AvenueSizeX = pagelayout(CEL.visPLOAvenueSizeX);
        public readonly static SRC AvenueSizeY = pagelayout(CEL.visPLOAvenueSizeY);
        public readonly static SRC BlockSizeX = pagelayout(CEL.visPLOBlockSizeX);
        public readonly static SRC BlockSizeY = pagelayout(CEL.visPLOBlockSizeY);
        public readonly static SRC CtrlAsInput = pagelayout(CEL.visPLOCtrlAsInput);
        public readonly static SRC DynamicsOff = pagelayout(CEL.visPLODynamicsOff);
        public readonly static SRC EnableGrid = pagelayout(CEL.visPLOEnableGrid);
        public readonly static SRC LineAdjustFrom = pagelayout(CEL.visPLOLineAdjustFrom);
        public readonly static SRC LineAdjustTo = pagelayout(CEL.visPLOLineAdjustTo);
        public readonly static SRC LineJumpCode = pagelayout(CEL.visPLOJumpCode);
        public readonly static SRC LineJumpFactorX = pagelayout(CEL.visPLOJumpFactorX);
        public readonly static SRC LineJumpFactorY = pagelayout(CEL.visPLOJumpFactorY);
        public readonly static SRC LineJumpStyle = pagelayout(CEL.visPLOJumpStyle);
        public readonly static SRC LineRouteExt = pagelayout(CEL.visPLOLineRouteExt);
        public readonly static SRC LineToLineX = pagelayout(CEL.visPLOLineToLineX);
        public readonly static SRC LineToLineY = pagelayout(CEL.visPLOLineToLineY);
        public readonly static SRC LineToNodeX = pagelayout(CEL.visPLOLineToNodeX);
        public readonly static SRC LineToNodeY = pagelayout(CEL.visPLOLineToNodeY);
        public readonly static SRC PageLineJumpDirX = pagelayout(CEL.visPLOJumpDirX);
        public readonly static SRC PageLineJumpDirY = pagelayout(CEL.visPLOJumpDirY);
        public readonly static SRC PageShapeSplit = pagelayout(CEL.visPLOSplit);
        public readonly static SRC PlaceDepth = pagelayout(CEL.visPLOPlaceDepth);
        public readonly static SRC PlaceFlip = pagelayout(CEL.visPLOPlaceFlip);
        public readonly static SRC PlaceStyle = pagelayout(CEL.visPLOPlaceStyle);
        public readonly static SRC PlowCode = pagelayout(CEL.visPLOPlowCode);
        public readonly static SRC ResizePage = pagelayout(CEL.visPLOResizePage);
        public readonly static SRC RouteStyle = pagelayout(CEL.visPLORouteStyle);


        // print properties

        public readonly static SRC PageLeftMargin =printprops (CEL.visPrintPropertiesLeftMargin);
        public readonly static SRC CenterX =printprops (CEL.visPrintPropertiesCenterX);
        public readonly static SRC CenterY =printprops (CEL.visPrintPropertiesCenterY);
        public readonly static SRC OnPage =printprops (CEL.visPrintPropertiesOnPage);
        public readonly static SRC PageBottomMargin =printprops (CEL.visPrintPropertiesBottomMargin);
        public readonly static SRC PageRightMargin =printprops (CEL.visPrintPropertiesRightMargin);
        public readonly static SRC PagesX =printprops (CEL.visPrintPropertiesPagesX);
        public readonly static SRC PagesY =printprops (CEL.visPrintPropertiesPagesY);
        public readonly static SRC PageTopMargin =printprops (CEL.visPrintPropertiesTopMargin);
        public readonly static SRC PaperKind =printprops (CEL.visPrintPropertiesPaperKind);
        public readonly static SRC PrintGrid =printprops (CEL.visPrintPropertiesPrintGrid);
        public readonly static SRC PrintPageOrientation =printprops (CEL.visPrintPropertiesPageOrientation);
        public readonly static SRC ScaleX =printprops (CEL.visPrintPropertiesScaleX);
        public readonly static SRC ScaleY =printprops (CEL.visPrintPropertiesScaleY);
        public readonly static SRC PaperSource =printprops (CEL.visPrintPropertiesPaperSource);

        // page properties

        public readonly static SRC DrawingScale = page (CEL.visPageDrawingScale);
        public readonly static SRC DrawingScaleType = page (CEL.visPageDrawScaleType);
        public readonly static SRC DrawingSizeType = page (CEL.visPageDrawSizeType);
        public readonly static SRC InhibitSnap = page (CEL.visPageInhibitSnap);
        public readonly static SRC PageHeight = page (CEL.visPageHeight);
        public readonly static SRC PageScale = page (CEL.visPageScale);
        public readonly static SRC PageWidth = page (CEL.visPageWidth);
        public readonly static SRC ShdwObliqueAngle = page (CEL.visPageShdwObliqueAngle);
        public readonly static SRC ShdwOffsetX = page (CEL.visPageShdwOffsetX);
        public readonly static SRC ShdwOffsetY = page (CEL.visPageShdwOffsetY);
        public readonly static SRC ShdwScaleFactor = page (CEL.visPageShdwScaleFactor);
        public readonly static SRC ShdwType = page (CEL.visPageShdwType);
        public readonly static SRC UIVisibility = page (CEL.visPageUIVisibility);

        // paragraph
        public readonly static SRC Para_Bullet = para ( CEL.visBulletIndex);
        public readonly static SRC Para_BulletFont = para ( CEL.visBulletFont);
        public readonly static SRC Para_BulletFontSize = para ( CEL.visBulletFontSize);
        public readonly static SRC Para_BulletStr = para ( CEL.visBulletString);
        public readonly static SRC Para_Flags = para ( CEL.visFlags);
        public readonly static SRC Para_HorzAlign = para ( CEL.visHorzAlign);
        public readonly static SRC Para_IndFirst = para ( CEL.visIndentFirst);
        public readonly static SRC Para_IndLeft = para ( CEL.visIndentLeft);
        public readonly static SRC Para_IndRight = para ( CEL.visIndentRight);
        public readonly static SRC Para_LocalizeBulletFont = para ( CEL.visLocalizeBulletFont);
        public readonly static SRC Para_SpAfter = para ( CEL.visSpaceAfter);
        public readonly static SRC Para_SpBefore = para ( CEL.visSpaceBefore);
        public readonly static SRC Para_SpLine = para ( CEL.visSpaceLine);
        public readonly static SRC Para_TextPosAfterBullet = para ( CEL.visTextPosAfterBullet);

        // protection

        public readonly static SRC LockAspect = lock_ ( CEL.visLockAspect);
        public readonly static SRC LockBegin = lock_ ( CEL.visLockBegin);
        public readonly static SRC LockCalcWH = lock_ ( CEL.visLockCalcWH);
        public readonly static SRC LockCrop = lock_ ( CEL.visLockCrop);
        public readonly static SRC LockCustProp = lock_ ( CEL.visLockCustProp);
        public readonly static SRC LockDelete = lock_ ( CEL.visLockDelete);
        public readonly static SRC LockEnd = lock_ ( CEL.visLockEnd);
        public readonly static SRC LockFormat = lock_ ( CEL.visLockFormat);
        public readonly static SRC LockFromGroupFormat = lock_ ( CEL.visLockFromGroupFormat);
        public readonly static SRC LockGroup = lock_ ( CEL.visLockGroup);
        public readonly static SRC LockHeight = lock_ ( CEL.visLockHeight);
        public readonly static SRC LockMoveX = lock_ ( CEL.visLockMoveX);
        public readonly static SRC LockMoveY = lock_ ( CEL.visLockMoveY);
        public readonly static SRC LockRotate = lock_ ( CEL.visLockRotate);
        public readonly static SRC LockSelect = lock_ ( CEL.visLockSelect);
        public readonly static SRC LockTextEdit = lock_ ( CEL.visLockTextEdit);
        public readonly static SRC LockThemeColors = lock_ ( CEL.visLockThemeColors);
        public readonly static SRC LockThemeEffects = lock_ ( CEL.visLockThemeEffects);
        public readonly static SRC LockVtxEdit = lock_ ( CEL.visLockVtxEdit);
        public readonly static SRC LockWidth = lock_ ( CEL.visLockWidth);


        // ruler and grid

        public readonly static SRC XGridDensity = rulergrid ( CEL.visXGridDensity);
        public readonly static SRC XGridOrigin = rulergrid ( CEL.visXGridOrigin);
        public readonly static SRC XGridSpacing = rulergrid ( CEL.visXGridSpacing);
        public readonly static SRC XRulerDensity = rulergrid ( CEL.visXRulerDensity);
        public readonly static SRC XRulerOrigin = rulergrid ( CEL.visXRulerOrigin);
        public readonly static SRC YGridDensity = rulergrid ( CEL.visYGridDensity);
        public readonly static SRC YGridOrigin = rulergrid ( CEL.visYGridOrigin);
        public readonly static SRC YGridSpacing = rulergrid ( CEL.visYGridSpacing);
        public readonly static SRC YRulerDensity = rulergrid ( CEL.visYRulerDensity);
        public readonly static SRC YRulerOrigin = rulergrid ( CEL.visYRulerOrigin);


        // Shape Tranform

        public readonly static SRC Angle =xformout (CEL.visXFormAngle);
        public readonly static SRC FlipX =xformout (CEL.visXFormFlipX);
        public readonly static SRC FlipY =xformout (CEL.visXFormFlipY);
        public readonly static SRC Height =xformout (CEL.visXFormHeight);
        public readonly static SRC LocPinX =xformout (CEL.visXFormLocPinX);
        public readonly static SRC LocPinY =xformout (CEL.visXFormLocPinY);
        public readonly static SRC PinX =xformout (CEL.visXFormPinX);
        public readonly static SRC PinY =xformout (CEL.visXFormPinY);
        public readonly static SRC ResizeMode =xformout (CEL.visXFormResizeMode);
        public readonly static SRC Width =xformout (CEL.visXFormWidth);

        // reviewer

        public readonly static SRC Reviewer_Color =    reviewer(  CEL.visReviewerColor);
        public readonly static SRC Reviewer_Initials = reviewer(  CEL.visReviewerInitials);
        public readonly static SRC Reviewer_Name =     reviewer(  CEL.visReviewerName);

        // shape data

        public readonly static SRC Prop_SortKey = prop (CEL.visCustPropsSortKey);
        public readonly static SRC Prop_Ask = prop (CEL.visCustPropsAsk);
        public readonly static SRC Prop_Calendar = prop (CEL.visCustPropsCalendar);
        public readonly static SRC Prop_Format = prop (CEL.visCustPropsFormat);
        public readonly static SRC Prop_Invisible = prop (CEL.visCustPropsInvis);
        public readonly static SRC Prop_Label = prop (CEL.visCustPropsLabel);
        public readonly static SRC Prop_LangID = prop (CEL.visCustPropsLangID);
        public readonly static SRC Prop_Prompt = prop (CEL.visCustPropsPrompt);
        public readonly static SRC Prop_Type = prop (CEL.visCustPropsType);
        public readonly static SRC Prop_Value = prop (CEL.visCustPropsValue);

        // Layers

        public readonly static SRC Layers_Active = layer(CEL.visLayerActive);
        public readonly static SRC Layers_Color = layer(CEL.visLayerColor);
        public readonly static SRC Layers_Glue = layer(CEL.visLayerGlue);
        public readonly static SRC Layers_Locked = layer(CEL.visLayerLock);
        public readonly static SRC Layers_Print = layer(CEL.visDocPreviewScope);
        public readonly static SRC Layers_Snap = layer(CEL.visLayerSnap);
        public readonly static SRC Layers_ColorTrans = layer(CEL.visLayerColorTrans);
        public readonly static SRC Layers_Visible = layer(CEL.visLayerVisible);


        //text transform
        public readonly static SRC TxtAngle = textxfrm ( CEL.visXFormAngle);
        public readonly static SRC TxtHeight = textxfrm ( CEL.visXFormHeight);
        public readonly static SRC TxtLocPinX = textxfrm ( CEL.visXFormLocPinX);
        public readonly static SRC TxtLocPinY = textxfrm ( CEL.visXFormLocPinY);
        public readonly static SRC TxtPinX = textxfrm ( CEL.visXFormPinX);
        public readonly static SRC TxtPinY = textxfrm ( CEL.visXFormPinY);
        public readonly static SRC TxtWidth = textxfrm ( CEL.visXFormWidth);


        // user defined cells
        public readonly static SRC User_Prompt = user ( CEL.visUserPrompt);
        public readonly static SRC User_Value =  user ( CEL.visUserValue);


        // Fields
        public readonly static SRC Fields_Calendar = field( CEL.visFieldCalendar);
        public readonly static SRC Fields_Format = field( CEL.visFieldFormat);
        public readonly static SRC Fields_ObjectKind = field( CEL.visFieldObjectKind);
        public readonly static SRC Fields_Type = field( CEL.visFieldType);
        public readonly static SRC Fields_UICat = field( CEL.visFieldUICategory);
        public readonly static SRC Fields_UICod = field( CEL.visFieldUICode);
        public readonly static SRC Fields_UIFmt = field( CEL.visFieldUIFormat);
        public readonly static SRC Fields_Value = field( CEL.visFieldCell);


        // text block format

        public readonly static SRC BottomMargin = text (CEL.visTxtBlkBottomMargin);
        public readonly static SRC DefaultTabStop = text (CEL.visTxtBlkDefaultTabStop);
        public readonly static SRC LeftMargin = text (CEL.visTxtBlkLeftMargin);
        public readonly static SRC RightMargin = text (CEL.visTxtBlkRightMargin);
        public readonly static SRC TextBkgnd = text (CEL.visTxtBlkBkgnd);
        public readonly static SRC TextBkgndTrans = text (CEL.visTxtBlkBkgndTrans);
        public readonly static SRC TextDirection = text (CEL.visTxtBlkDirection);
        public readonly static SRC TopMargin = text (CEL.visTxtBlkTopMargin);
        public readonly static SRC VerticalAlign = text (CEL.visTxtBlkVerticalAlign);

        // Action tags
        public readonly static SRC SmartTags_ButtonFace =smarttag ( CEL.visSmartTagButtonFace);
        public readonly static SRC SmartTags_Description =smarttag ( CEL.visSmartTagDescription);
        public readonly static SRC SmartTags_Disabled =smarttag ( CEL.visSmartTagDisabled);
        public readonly static SRC SmartTags_DisplayMode =smarttag ( CEL.visSmartTagDisplayMode);
        public readonly static SRC SmartTags_TagName =smarttag ( CEL.visSmartTagName);
        public readonly static SRC SmartTags_X =smarttag ( CEL.visSmartTagX);
        public readonly static SRC SmartTags_XJustify =smarttag ( CEL.visSmartTagXJustify);
        public readonly static SRC SmartTags_Y =smarttag ( CEL.visSmartTagY);
        public readonly static SRC SmartTags_YJustify =smarttag ( CEL.visSmartTagYJustify);

        // style
        public readonly static SRC EnableFillProps = style(CEL.visStyleIncludesFill);
        public readonly static SRC EnableLineProps = style(CEL.visStyleIncludesLine);
        public readonly static SRC EnableTextProps = style(CEL.visStyleIncludesText);
        public readonly static SRC HideText = style(CEL.visStyleHidden);


        //tabs
        public readonly static SRC Tabs_Alignment = tab( CEL.visTabAlign);
        public readonly static SRC Tabs_Position =  tab( CEL.visTabPos);
        public readonly static SRC Tabs_StopCount = tab( CEL.visTabStopCount);

        // shape layout
        public readonly static SRC ConFixedCode = shapelayout( CEL.visSLOConFixedCode);
        public readonly static SRC ConLineJumpCode = shapelayout( CEL.visSLOJumpCode);
        public readonly static SRC ConLineJumpDirX = shapelayout( CEL.visSLOJumpDirX);
        public readonly static SRC ConLineJumpDirY = shapelayout( CEL.visSLOJumpDirY);
        public readonly static SRC ConLineJumpStyle = shapelayout( CEL.visSLOJumpStyle);
        public readonly static SRC ConLineRouteExt = shapelayout( CEL.visSLOLineRouteExt);
        public readonly static SRC ShapeFixedCode = shapelayout( CEL.visSLOFixedCode);
        public readonly static SRC ShapePermeablePlace = shapelayout( CEL.visSLOPermeablePlace);
        public readonly static SRC ShapePermeableX = shapelayout( CEL.visSLOPermX);
        public readonly static SRC ShapePermeableY = shapelayout( CEL.visSLOPermY);
        public readonly static SRC ShapePlaceFlip = shapelayout( CEL.visSLOPlaceFlip);
        public readonly static SRC ShapePlaceStyle = shapelayout( CEL.visSLOPlaceStyle);
        public readonly static SRC ShapePlowCode = shapelayout( CEL.visSLOPlowCode);
        public readonly static SRC ShapeRouteStyle = shapelayout( CEL.visSLORouteStyle);
        public readonly static SRC ShapeSplit = shapelayout( CEL.visSLOSplit);
        public readonly static SRC ShapeSplittable = shapelayout( CEL.visSLOSplittable);

        // Static Methods

        public static Dictionary<string, VA.ShapeSheet.SRC> GetSRCDictionary()
        {
            var fields = GetSRCFields();

            var fields_name_to_value = new Dictionary<string, VA.ShapeSheet.SRC>();
            foreach (var field in fields)
            {
                fields_name_to_value[field.Name] = (VA.ShapeSheet.SRC)field.GetValue(null);
            }

            return fields_name_to_value;
        }

        private static List<FieldInfo> GetSRCFields()
        {
            var srcconstants_t = typeof(VA.ShapeSheet.SRCConstants);
            var fields = srcconstants_t.GetFields()
                .Where(m => m.FieldType == typeof(VA.ShapeSheet.SRC))
                .Where(m => m.IsPublic)
                .Where(m => m.IsStatic)
                .ToList();
            return fields;
        }
    }
}