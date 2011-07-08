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
        // Actions
        public readonly static SRC Actions_Action = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionAction);
        public readonly static SRC Actions_BeginGroup = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionBeginGroup);
        public readonly static SRC Actions_ButtonFace = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionButtonFace);
        public readonly static SRC Actions_Checked = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionChecked);
        public readonly static SRC Actions_Disabled = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionDisabled);
        public readonly static SRC Actions_Invisible = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionInvisible);
        public readonly static SRC Actions_Menu = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionMenu);
        public readonly static SRC Actions_ReadOnly = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionReadOnly);
        public readonly static SRC Actions_SortKey = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionSortKey);
        public readonly static SRC Actions_TagName = new SRC(SEC.visSectionAction, ROW.visRowAction, CEL.visActionTagName);

        // Alignment
        public readonly static SRC AlignBottom = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignBottom);
        public readonly static SRC AlignCenter = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignCenter);
        public readonly static SRC AlignLeft = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignLeft);
        public readonly static SRC AlignMiddle = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignMiddle);
        public readonly static SRC AlignRight = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignRight);
        public readonly static SRC AlignTop = new SRC(SEC.visSectionObject, ROW.visRowAlign, CEL.visAlignTop);

        // Annotation
        public readonly static SRC Annotation_Comment = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationComment);
        public readonly static SRC Annotation_Date = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationDate);
        public readonly static SRC Annotation_LangID = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationLangID);
        public readonly static SRC Annotation_MarkerIndex = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationMarkerIndex);
        public readonly static SRC Annotation_X = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationX);
        public readonly static SRC Annotation_Y = new SRC(SEC.visSectionAnnotation, ROW.visRowAnnotation, CEL.visAnnotationY);

        // Character
        public readonly static SRC Char_AsianFont = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterAsianFont);
        public readonly static SRC Char_Case = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterCase);
        public readonly static SRC Char_Color = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColor);
        public readonly static SRC Char_ComplexScriptFont = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptFont);
        public readonly static SRC Char_ComplexScriptSize = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterComplexScriptSize);
        public readonly static SRC Char_DoubleStrikethrough = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDoubleStrikethrough);
        public readonly static SRC Char_DblUnderline = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterDblUnderline);
        public readonly static SRC Char_Font = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFont);
        public readonly static SRC Char_LangID = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLangID);
        public readonly static SRC Char_Locale = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocale);
        public readonly static SRC Char_LocalizeFont = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLocalizeFont);
        public readonly static SRC Char_Overline = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterOverline);
        public readonly static SRC Char_Perpendicular = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPerpendicular);
        public readonly static SRC Char_Pos = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterPos);
        public readonly static SRC RTLText = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterRTLText);
        public readonly static SRC Char_FontScale = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterFontScale);
        public readonly static SRC Char_Size = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterSize);
        public readonly static SRC Char_Letterspace = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterLetterspace);
        public readonly static SRC Char_Strikethru = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStrikethru);
        public readonly static SRC Char_Style = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterStyle);
        public readonly static SRC Char_ColorTrans = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterColorTrans);
        
        public readonly static SRC UseVertical = new SRC(SEC.visSectionCharacter, ROW.visRowCharacter, CEL.visCharacterUseVertical);

        // Connections
        public readonly static SRC Connections_D = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctD);
        public readonly static SRC Connections_DirX = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirX);
        public readonly static SRC Connections_DirY = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctDirY);
        public readonly static SRC Connections_Type = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visCnnctType);
        public readonly static SRC Connections_X = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visX);
        public readonly static SRC Connections_Y = new SRC(SEC.visSectionConnectionPts, ROW.visRowConnectionPts, CEL.visY);

        // Controls
        public readonly static SRC Controls_CanGlue = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlGlue);
        public readonly static SRC Controls_Tip = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlTip);
        public readonly static SRC Controls_XCon = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXCon);
        public readonly static SRC Controls_X = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlX);
        public readonly static SRC Controls_XDyn = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlXDyn);
        public readonly static SRC Controls_YCon = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYCon);
        public readonly static SRC Controls_Y = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlY);
        public readonly static SRC Controls_YDyn = new SRC(SEC.visSectionControls, ROW.visRowControl, CEL.visCtlYDyn);

        // Document Properties

        public readonly static SRC AddMarkup = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocAddMarkup);
        public readonly static SRC DocLangID = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLangID);
        public readonly static SRC LockPreview = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocLockPreview);
        public readonly static SRC OutputFormat = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocOutputFormat);
        public readonly static SRC PreviewQuality = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewQuality);
        public readonly static SRC PreviewScope = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocPreviewScope);
        public readonly static SRC ViewMarkup = new SRC(SEC.visSectionObject, ROW.visRowDoc, CEL.visDocViewMarkup);


        // Events
        public readonly static SRC EventDblClick = new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDblClick);
        public readonly static SRC EventDrop = new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellDrop);
        public readonly static SRC EventMultiDrop = new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellMultiDrop);
        public readonly static SRC EventXFMod = new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellXFMod);
        public readonly static SRC TheText = new SRC(SEC.visSectionObject, ROW.visRowEvent, CEL.visEvtCellTheText);

        // ForeignImageInfo
        public readonly static SRC ImgHeight = new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgHeight);
        public readonly static SRC ImgOffsetX = new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetX);
        public readonly static SRC ImgOffsetY = new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgOffsetY);
        public readonly static SRC ImgWidth = new SRC(SEC.visSectionObject, ROW.visRowForeign, CEL.visFrgnImgWidth);


        // Geometry
        public readonly static SRC Geometry_A = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visBow);
        public readonly static SRC Geometry_B = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visControlX);
        public readonly static SRC Geometry_C = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visEccentricityAngle);
        public readonly static SRC Geometry_D = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visAspectRatio);
        public readonly static SRC Geometry_E = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visNURBSData);
        public readonly static SRC Geometry_NoFill = new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoFill);
        public readonly static SRC Geometry_NoLine = new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoLine);
        public readonly static SRC Geometry_NoShow = new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoShow);
        public readonly static SRC Geometry_NoSnap = new SRC(SEC.visSectionFirstComponent, ROW.visRowComponent, CEL.visCompNoSnap);
        public readonly static SRC Geometry_X = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visX);
        public readonly static SRC Geometry_Y = new SRC(SEC.visSectionFirstComponent, ROW.visRowVertex, CEL.visY);


        // Fill Format

        public readonly static SRC FillBkgnd = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgnd);
        public readonly static SRC FillBkgndTrans = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillBkgndTrans);
        public readonly static SRC FillForegnd = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd);
        public readonly static SRC FillForegndTrans = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegndTrans);
        public readonly static SRC FillPattern = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillPattern);
        public readonly static SRC ShapeShdwObliqueAngle = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwObliqueAngle);
        public readonly static SRC ShapeShdwOffsetX = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetX);
        public readonly static SRC ShapeShdwOffsetY = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwOffsetY);
        public readonly static SRC ShapeShdwScaleFactor = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwScaleFactor);
        public readonly static SRC ShapeShdwType = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwType);
        public readonly static SRC ShdwBkgnd = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgnd);
        public readonly static SRC ShdwBkgndTrans = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwBkgndTrans);
        public readonly static SRC ShdwForegnd = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegnd);
        public readonly static SRC ShdwForegndTrans = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwForegndTrans);
        public readonly static SRC ShdwPattern = new SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillShdwPattern);

        // GlueInfo
        public readonly static SRC BegTrigger = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visBegTrigger);
        public readonly static SRC EndTrigger = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visEndTrigger);
        public readonly static SRC GlueType = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visGlueType);
        public readonly static SRC WalkPreference = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visWalkPref);

        // GroupProperties

        public readonly static SRC DisplayMode = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDisplayMode);
        public readonly static SRC DontMoveChildren = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupDontMoveChildren);
        public readonly static SRC IsDropTarget = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsDropTarget);
        public readonly static SRC IsSnapTarget = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsSnapTarget);
        public readonly static SRC IsTextEditTarget = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupIsTextEditTarget);
        public readonly static SRC SelectMode = new SRC(SEC.visSectionObject, ROW.visRowGroup, CEL.visGroupSelectMode);

        // Hyperlinks

        public readonly static SRC Hyperlink_Address = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkAddress);
        public readonly static SRC Hyperlink_Default = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDefault);
        public readonly static SRC Hyperlink_Description = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkDescription);
        public readonly static SRC Hyperlink_ExtraInfo = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkExtraInfo);
        public readonly static SRC Hyperlink_Frame = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkFrame);
        public readonly static SRC Hyperlink_Invisible = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkInvisible);
        public readonly static SRC Hyperlink_NewWindow = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkNewWin);
        public readonly static SRC Hyperlink_SortKey = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSortKey);
        public readonly static SRC Hyperlink_SubAddress = new SRC(SEC.visSectionHyperlink, ROW.visRow1stHyperlink, CEL.visHLinkSubAddress);

        // Image Properties
        public readonly static SRC Blur = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBlur);
        public readonly static SRC Brightness = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageBrightness);
        public readonly static SRC Contrast = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageContrast);
        public readonly static SRC Denoise = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageDenoise);
        public readonly static SRC Gamma = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageGamma);
        public readonly static SRC Sharpen = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageSharpen);
        public readonly static SRC Transparency = new SRC(SEC.visSectionObject, ROW.visRowImage, CEL.visImageTransparency);


        // Line format
        public readonly static SRC BeginArrow = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrow);
        public readonly static SRC BeginArrowSize = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineBeginArrowSize);
        public readonly static SRC EndArrow = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrow);
        public readonly static SRC EndArrowSize = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndArrowSize);
        public readonly static SRC LineCap = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineEndCap);
        public readonly static SRC LineColor = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColor);
        public readonly static SRC LineColorTrans = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineColorTrans);
        public readonly static SRC LinePattern = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLinePattern);
        public readonly static SRC LineWeight = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineWeight);
        public readonly static SRC Rounding = new SRC(SEC.visSectionObject, ROW.visRowLine, CEL.visLineRounding);


        // Miscellaneous

        public readonly static SRC Calendar = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjCalendar);
        public readonly static SRC Comment = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visComment);
        public readonly static SRC DropOnPageScale = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjDropOnPageScale);
        public readonly static SRC DynFeedback = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDynFeedback);
        public readonly static SRC IsDropSource = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visDropSource);
        public readonly static SRC LangID = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLangID);
        public readonly static SRC LocalizeMerge = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visObjLocalizeMerge);
        public readonly static SRC NoAlignBox = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoAlignBox);
        public readonly static SRC NoCtlHandles = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoCtlHandles);
        public readonly static SRC NoLiveDynamics = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoLiveDynamics);
        public readonly static SRC NonPrinting = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNonPrinting);
        public readonly static SRC NoObjHandles = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visNoObjHandles);
        public readonly static SRC ObjType = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visLOFlags);
        public readonly static SRC UpdateAlignBox = new SRC(SEC.visSectionObject, ROW.visRowMisc, CEL.visUpdateAlignBox);

        // 1d endpoints

        public readonly static SRC BeginX = new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginX);
        public readonly static SRC BeginY = new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DBeginY);
        public readonly static SRC EndX = new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndX);
        public readonly static SRC EndY = new SRC(SEC.visSectionObject, ROW.visRowXForm1D, CEL.vis1DEndY);


        // page layout
        public readonly static SRC AvenueSizeX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeX);
        public readonly static SRC AvenueSizeY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOAvenueSizeY);
        public readonly static SRC BlockSizeX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeX);
        public readonly static SRC BlockSizeY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOBlockSizeY);
        public readonly static SRC CtrlAsInput = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOCtrlAsInput);
        public readonly static SRC DynamicsOff = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLODynamicsOff);
        public readonly static SRC EnableGrid = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOEnableGrid);
        public readonly static SRC LineAdjustFrom = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustFrom);
        public readonly static SRC LineAdjustTo = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineAdjustTo);
        public readonly static SRC LineJumpCode = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpCode);
        public readonly static SRC LineJumpFactorX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorX);
        public readonly static SRC LineJumpFactorY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpFactorY);
        public readonly static SRC LineJumpStyle = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpStyle);
        public readonly static SRC LineRouteExt = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineRouteExt);
        public readonly static SRC LineToLineX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineX);
        public readonly static SRC LineToLineY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToLineY);
        public readonly static SRC LineToNodeX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeX);
        public readonly static SRC LineToNodeY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOLineToNodeY);
        public readonly static SRC PageLineJumpDirX = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirX);
        public readonly static SRC PageLineJumpDirY = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOJumpDirY);
        public readonly static SRC PageShapeSplit = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOSplit);
        public readonly static SRC PlaceDepth = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceDepth);
        public readonly static SRC PlaceFlip = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceFlip);
        public readonly static SRC PlaceStyle = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlaceStyle);
        public readonly static SRC PlowCode = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOPlowCode);
        public readonly static SRC ResizePage = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLOResizePage);
        public readonly static SRC RouteStyle = new SRC(SEC.visSectionObject, ROW.visRowPageLayout, CEL.visPLORouteStyle);


        // print properties

        public readonly static SRC PageLeftMargin = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesLeftMargin);
        public readonly static SRC CenterX = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterX);
        public readonly static SRC CenterY = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesCenterY);
        public readonly static SRC OnPage = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesOnPage);
        public readonly static SRC PageBottomMargin = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesBottomMargin);
        public readonly static SRC PageRightMargin = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesRightMargin);
        public readonly static SRC PagesX = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesX);
        public readonly static SRC PagesY = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPagesY);
        public readonly static SRC PageTopMargin = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesTopMargin);
        public readonly static SRC PaperKind = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperKind);
        public readonly static SRC PrintGrid = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPrintGrid);
        public readonly static SRC PrintPageOrientation = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPageOrientation);
        public readonly static SRC ScaleX = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleX);
        public readonly static SRC ScaleY = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesScaleY);
        public readonly static SRC PaperSource = new SRC(SEC.visSectionObject, ROW.visRowPrintProperties, CEL.visPrintPropertiesPaperSource);

        // page properties

        public readonly static SRC DrawingScale = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawingScale);
        public readonly static SRC DrawingScaleType = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawScaleType);
        public readonly static SRC DrawingSizeType = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageDrawSizeType);
        public readonly static SRC InhibitSnap = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageInhibitSnap);
        public readonly static SRC PageHeight = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageHeight);
        public readonly static SRC PageScale = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageScale);
        public readonly static SRC PageWidth = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageWidth);
        public readonly static SRC ShdwObliqueAngle = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwObliqueAngle);
        public readonly static SRC ShdwOffsetX = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetX);
        public readonly static SRC ShdwOffsetY = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwOffsetY);
        public readonly static SRC ShdwScaleFactor = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwScaleFactor);
        public readonly static SRC ShdwType = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageShdwType);
        public readonly static SRC UIVisibility = new SRC(SEC.visSectionObject, ROW.visRowPage, CEL.visPageUIVisibility);

        // paragraph
        public readonly static SRC Para_Bullet = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletIndex);
        public readonly static SRC Para_BulletFont = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFont);
        public readonly static SRC Para_BulletFontSize = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletFontSize);
        public readonly static SRC Para_BulletStr = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visBulletString);
        public readonly static SRC Para_Flags = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visFlags);
        public readonly static SRC Para_HorzAlign = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visHorzAlign);
        public readonly static SRC Para_IndFirst = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentFirst);
        public readonly static SRC Para_IndLeft = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentLeft);
        public readonly static SRC Para_IndRight = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visIndentRight);
        public readonly static SRC Para_LocalizeBulletFont = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visLocalizeBulletFont);
        public readonly static SRC Para_SpAfter = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceAfter);
        public readonly static SRC Para_SpBefore = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceBefore);
        public readonly static SRC Para_SpLine = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visSpaceLine);
        public readonly static SRC Para_TextPosAfterBullet = new SRC(SEC.visSectionParagraph, ROW.visRowParagraph, CEL.visTextPosAfterBullet);

        // protection

        public readonly static SRC LockAspect = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockAspect);
        public readonly static SRC LockBegin = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockBegin);
        public readonly static SRC LockCalcWH = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCalcWH);
        public readonly static SRC LockCrop = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCrop);
        public readonly static SRC LockCustProp = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockCustProp);
        public readonly static SRC LockDelete = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockDelete);
        public readonly static SRC LockEnd = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockEnd);
        public readonly static SRC LockFormat = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFormat);
        public readonly static SRC LockFromGroupFormat = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockFromGroupFormat);
        public readonly static SRC LockGroup = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockGroup);
        public readonly static SRC LockHeight = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockHeight);
        public readonly static SRC LockMoveX = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveX);
        public readonly static SRC LockMoveY = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockMoveY);
        public readonly static SRC LockRotate = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockRotate);
        public readonly static SRC LockSelect = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockSelect);
        public readonly static SRC LockTextEdit = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockTextEdit);
        public readonly static SRC LockThemeColors = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeColors);
        public readonly static SRC LockThemeEffects = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockThemeEffects);
        public readonly static SRC LockVtxEdit = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockVtxEdit);
        public readonly static SRC LockWidth = new SRC(SEC.visSectionObject, ROW.visRowLock, CEL.visLockWidth);


        // ruler and grid

        public readonly static SRC XGridDensity = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridDensity);
        public readonly static SRC XGridOrigin = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridOrigin);
        public readonly static SRC XGridSpacing = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXGridSpacing);
        public readonly static SRC XRulerDensity = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerDensity);
        public readonly static SRC XRulerOrigin = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visXRulerOrigin);
        public readonly static SRC YGridDensity = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridDensity);
        public readonly static SRC YGridOrigin = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridOrigin);
        public readonly static SRC YGridSpacing = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYGridSpacing);
        public readonly static SRC YRulerDensity = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerDensity);
        public readonly static SRC YRulerOrigin = new SRC(SEC.visSectionObject, ROW.visRowRulerGrid, CEL.visYRulerOrigin);


        // Shape Tranform

        public readonly static SRC Angle = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormAngle);
        public readonly static SRC FlipX = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipX);
        public readonly static SRC FlipY = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormFlipY);
        public readonly static SRC Height = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormHeight);
        public readonly static SRC LocPinX = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinX);
        public readonly static SRC LocPinY = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormLocPinY);
        public readonly static SRC PinX = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinX);
        public readonly static SRC PinY = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormPinY);
        public readonly static SRC ResizeMode = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormResizeMode);
        public readonly static SRC Width = new SRC(SEC.visSectionObject, ROW.visRowXFormOut, CEL.visXFormWidth);

        // reviewer

        public readonly static SRC Reviewer_Color = new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerColor);
        public readonly static SRC Reviewer_Initials = new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerInitials);
        public readonly static SRC Reviewer_Name = new SRC(SEC.visSectionReviewer, ROW.visRowReviewer, CEL.visReviewerName);

        // shape data

        public readonly static SRC Prop_SortKey = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsSortKey);
        public readonly static SRC Prop_Ask = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsAsk);
        public readonly static SRC Prop_Calendar = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsCalendar);
        public readonly static SRC Prop_Format = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsFormat);
        public readonly static SRC Prop_Invisible = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsInvis);
        public readonly static SRC Prop_Label = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLabel);
        public readonly static SRC Prop_LangID = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsLangID);
        public readonly static SRC Prop_Prompt = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsPrompt);
        public readonly static SRC Prop_Type = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsType);
        public readonly static SRC Prop_Value = new SRC(SEC.visSectionProp, ROW.visRowProp, CEL.visCustPropsValue);

        // Layers

        public readonly static SRC Layers_Active = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerActive);
        public readonly static SRC Layers_Color = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColor);
        public readonly static SRC Layers_Glue = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerGlue);
        public readonly static SRC Layers_Locked = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerLock);
        public readonly static SRC Layers_Print = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visDocPreviewScope);
        public readonly static SRC Layers_Snap = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerSnap);
        public readonly static SRC Layers_ColorTrans = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerColorTrans);
        public readonly static SRC Layers_Visible = new SRC(SEC.visSectionLayer, ROW.visRowLayer, CEL.visLayerVisible);


        //text transform
        public readonly static SRC TxtAngle = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormAngle);
        public readonly static SRC TxtHeight = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormHeight);
        public readonly static SRC TxtLocPinX = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinX);
        public readonly static SRC TxtLocPinY = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormLocPinY);
        public readonly static SRC TxtPinX = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinX);
        public readonly static SRC TxtPinY = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormPinY);
        public readonly static SRC TxtWidth = new SRC(SEC.visSectionObject, ROW.visRowTextXForm, CEL.visXFormWidth);


        // user defined cells
        public readonly static SRC User_Prompt = new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserPrompt);
        public readonly static SRC User_Value = new SRC(SEC.visSectionUser, ROW.visRowUser, CEL.visUserValue);


        // Fields
        public readonly static SRC Fields_Calendar = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCalendar);
        public readonly static SRC Fields_Format = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldFormat);
        public readonly static SRC Fields_ObjectKind = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldObjectKind);
        public readonly static SRC Fields_Type = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldType);
        public readonly static SRC Fields_UICat = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICategory);
        public readonly static SRC Fields_UICod = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUICode);
        public readonly static SRC Fields_UIFmt = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldUIFormat);
        public readonly static SRC Fields_Value = new SRC(SEC.visSectionTextField, ROW.visRowField, CEL.visFieldCell);


        // text block format

        public readonly static SRC BottomMargin = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBottomMargin);
        public readonly static SRC DefaultTabStop = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDefaultTabStop);
        public readonly static SRC LeftMargin = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkLeftMargin);
        public readonly static SRC RightMargin = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkRightMargin);
        public readonly static SRC TextBkgnd = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgnd);
        public readonly static SRC TextBkgndTrans = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkBkgndTrans);
        public readonly static SRC TextDirection = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkDirection);
        public readonly static SRC TopMargin = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkTopMargin);
        public readonly static SRC VerticalAlign = new SRC(SEC.visSectionObject, ROW.visRowText, CEL.visTxtBlkVerticalAlign);

        // Action tags
        public readonly static SRC SmartTags_ButtonFace = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagButtonFace);
        public readonly static SRC SmartTags_Description = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDescription);
        public readonly static SRC SmartTags_Disabled = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisabled);
        public readonly static SRC SmartTags_DisplayMode = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagDisplayMode);
        public readonly static SRC SmartTags_TagName = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagName);
        public readonly static SRC SmartTags_X = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagX);
        public readonly static SRC SmartTags_XJustify = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagXJustify);
        public readonly static SRC SmartTags_Y = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagY);
        public readonly static SRC SmartTags_YJustify = new SRC(SEC.visSectionSmartTag, ROW.visRowSmartTag, CEL.visSmartTagYJustify);

        // style
        public readonly static SRC EnableFillProps = new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesFill);
        public readonly static SRC EnableLineProps = new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesLine);
        public readonly static SRC EnableTextProps = new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleIncludesText);
        public readonly static SRC HideText = new SRC(SEC.visSectionObject, ROW.visRowStyle, CEL.visStyleHidden);


        //tabs
        public readonly static SRC Tabs_Alignment = new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabAlign);
        public readonly static SRC Tabs_Position = new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabPos);
        public readonly static SRC Tabs_StopCount = new SRC(SEC.visSectionTab, ROW.visRowTab, CEL.visTabStopCount);

        // shape layout
        public readonly static SRC ConFixedCode = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOConFixedCode);
        public readonly static SRC ConLineJumpCode = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpCode);
        public readonly static SRC ConLineJumpDirX = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirX);
        public readonly static SRC ConLineJumpDirY = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpDirY);
        public readonly static SRC ConLineJumpStyle = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOJumpStyle);
        public readonly static SRC ConLineRouteExt = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOLineRouteExt);
        public readonly static SRC ShapeFixedCode = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOFixedCode);
        public readonly static SRC ShapePermeablePlace = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermeablePlace);
        public readonly static SRC ShapePermeableX = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermX);
        public readonly static SRC ShapePermeableY = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPermY);
        public readonly static SRC ShapePlaceFlip = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceFlip);
        public readonly static SRC ShapePlaceStyle = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlaceStyle);
        public readonly static SRC ShapePlowCode = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOPlowCode);
        public readonly static SRC ShapeRouteStyle = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLORouteStyle);
        public readonly static SRC ShapeSplit = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplit);
        public readonly static SRC ShapeSplittable = new SRC(SEC.visSectionObject, ROW.visRowShapeLayout, CEL.visSLOSplittable);

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