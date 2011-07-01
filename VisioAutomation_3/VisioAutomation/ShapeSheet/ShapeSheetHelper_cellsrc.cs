using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static partial class ShapeSheetHelper
    {
        public static string TryGetNameFromSRC(VA.ShapeSheet.SRC src)
        {
            switch (src.Section)
            {
                case ((short)IVisio.VisSectionIndices.visSectionObject):
                    {
                        return TryGetNameFromSRC_Section_Object(src);
                    }
                default:
                    break;
            }
            return null;
        }

        private static string TryGetNameFromSRC_Section_Object(VA.ShapeSheet.SRC src)
        {
            switch (src.Row)
            {
                case ((short)IVisio.VisRowIndices.visRowFill):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visFillBkgnd): return "FillBkgnd";
                            case ((short)IVisio.VisCellIndices.visFillBkgndTrans): return "FillBkgndTrans";
                            case ((short)IVisio.VisCellIndices.visFillForegnd): return "FillForegnd";
                            case ((short)IVisio.VisCellIndices.visFillForegndTrans): return "FillForegndTrans";
                            case ((short)IVisio.VisCellIndices.visFillPattern): return "FillPattern";
                            case ((short)IVisio.VisCellIndices.visFillShdwBkgnd): return "ShdwBkgnd";
                            case ((short)IVisio.VisCellIndices.visFillShdwBkgndTrans): return "ShdwBkgndTrans";
                            case ((short)IVisio.VisCellIndices.visFillShdwForegnd): return "ShdwForegnd";
                            case ((short)IVisio.VisCellIndices.visFillShdwForegndTrans): return "ShdwForegndTrans";
                            case ((short)IVisio.VisCellIndices.visFillShdwObliqueAngle): return "ShdwObliqueAngle";
                            case ((short)IVisio.VisCellIndices.visFillShdwOffsetX): return "ShdwOffsetX";
                            case ((short)IVisio.VisCellIndices.visFillShdwOffsetY): return "ShdwOffsetY";
                            case ((short)IVisio.VisCellIndices.visFillShdwPattern): return "ShdwPattern";
                            case ((short)IVisio.VisCellIndices.visFillShdwScaleFactor): return "ShdwScaleFactor";
                            case ((short)IVisio.VisCellIndices.visFillShdwType): return "FillShdwType";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowLine):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visLineBeginArrow): return "BeginArrow";
                            case ((short)IVisio.VisCellIndices.visLineBeginArrowSize): return "BeginArrowSize";
                            case ((short)IVisio.VisCellIndices.visLineColor): return "LineColor";
                            case ((short)IVisio.VisCellIndices.visLineColorTrans): return "LineColorTrans";
                            case ((short)IVisio.VisCellIndices.visLineEndArrow): return "LineEndArrow";
                            case ((short)IVisio.VisCellIndices.visLineEndArrowSize): return "LineEndArrowSize";
                            case ((short)IVisio.VisCellIndices.visLineEndCap): return "LineEndCap";
                            case ((short)IVisio.VisCellIndices.visLinePattern): return "LinePattern";
                            case ((short)IVisio.VisCellIndices.visLineRounding): return "LineRounding";
                            case ((short)IVisio.VisCellIndices.visLineWeight): return "LineWeight";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowXFormOut):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visXFormAngle): return "Angle";
                            case ((short)IVisio.VisCellIndices.visXFormFlipX): return "FlipX";
                            case ((short)IVisio.VisCellIndices.visXFormFlipY): return "FlipY";
                            case ((short)IVisio.VisCellIndices.visXFormHeight): return "Height";
                            case ((short)IVisio.VisCellIndices.visXFormLocPinX): return "LocPinX";
                            case ((short)IVisio.VisCellIndices.visXFormLocPinY): return "LocPinY";
                            case ((short)IVisio.VisCellIndices.visXFormPinX): return "PinX";
                            case ((short)IVisio.VisCellIndices.visXFormPinY): return "PinY";
                            case ((short)IVisio.VisCellIndices.visXFormResizeMode): return "ResizeMode";
                            case ((short)IVisio.VisCellIndices.visXFormWidth): return "FormWidth";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowAlign):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visAlignBottom): return "AlignBottom";
                            case ((short)IVisio.VisCellIndices.visAlignCenter): return "AlignCenter";
                            case ((short)IVisio.VisCellIndices.visAlignLeft): return "AlignLeft";
                            case ((short)IVisio.VisCellIndices.visAlignRight): return "AlignRight";
                            case ((short)IVisio.VisCellIndices.visAlignTop): return "AlignTop";
                            case ((short)IVisio.VisCellIndices.visAlignMiddle): return "AlignMiddle";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowDoc):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visDocAddMarkup): return "AddMarkup";
                            case ((short)IVisio.VisCellIndices.visDocLangID): return "DocLangID";
                            case ((short)IVisio.VisCellIndices.visDocLockPreview): return "LockPreview";
                            case ((short)IVisio.VisCellIndices.visDocOutputFormat): return "OutputFormat";
                            case ((short)IVisio.VisCellIndices.visDocPreviewQuality): return "PreviewQuality";
                            case ((short)IVisio.VisCellIndices.visDocPreviewScope): return "PreviewScope";
                            case ((short)IVisio.VisCellIndices.visDocViewMarkup): return "PreviewScope";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRow1stHyperlink):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visHLinkDescription): return "Hyperlink.Description";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowForeign):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visFrgnImgHeight): return "ImgHeight";
                            case ((short)IVisio.VisCellIndices.visFrgnImgOffsetX): return "ImgOffsetX";
                            case ((short)IVisio.VisCellIndices.visFrgnImgOffsetY): return "ImgOffsetY";
                            case ((short)IVisio.VisCellIndices.visFrgnImgWidth): return "ImgWidth";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowEvent):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visEvtCellDblClick): return "EventDblClick";
                            case ((short)IVisio.VisCellIndices.visEvtCellDrop): return "EventDrop";
                            case ((short)IVisio.VisCellIndices.visEvtCellMultiDrop): return "EventMultiDrop";
                            case ((short)IVisio.VisCellIndices.visEvtCellXFMod): return "EventXFMod";
                            case ((short)IVisio.VisCellIndices.visEvtCellTheText): return "TheText";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowMisc):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visBegTrigger): return "BegTrigger";
                            case ((short)IVisio.VisCellIndices.visEndTrigger): return "EndTrigger";
                            case ((short)IVisio.VisCellIndices.visGlueType): return "GlueType";
                            case ((short)IVisio.VisCellIndices.visWalkPref): return "WalkPreference";
                            case ((short)IVisio.VisCellIndices.visObjCalendar): return "Calendar";
                            case ((short)IVisio.VisCellIndices.visComment): return "Comment";
                            case ((short)IVisio.VisCellIndices.visObjDropOnPageScale): return "DropOnPageScale";
                            case ((short)IVisio.VisCellIndices.visDynFeedback): return "DynFeedback";
                            case ((short)IVisio.VisCellIndices.visDropSource): return "IsDropSource";
                            case ((short)IVisio.VisCellIndices.visObjLangID): return "LangID";
                            case ((short)IVisio.VisCellIndices.visObjLocalizeMerge): return "LocalizeMerge";
                            case ((short)IVisio.VisCellIndices.visNoAlignBox): return "NoAlignBox";
                            case ((short)IVisio.VisCellIndices.visNoCtlHandles): return "NoCtlHandles";
                            case ((short)IVisio.VisCellIndices.visNoLiveDynamics): return "NoLiveDynamics";
                            case ((short)IVisio.VisCellIndices.visNonPrinting): return "NonPrinting";
                            case ((short)IVisio.VisCellIndices.visNoObjHandles): return "NoObjHandles";
                            case ((short)IVisio.VisCellIndices.visLOFlags): return "ObjType";
                            case ((short)IVisio.VisCellIndices.visUpdateAlignBox): return "UpdateAlignBox";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowGroup):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visGroupDisplayMode): return "DisplayMode";
                            case ((short)IVisio.VisCellIndices.visGroupDontMoveChildren): return "DontMoveChildren";
                            case ((short)IVisio.VisCellIndices.visGroupIsDropTarget): return "IsDropTarget";
                            case ((short)IVisio.VisCellIndices.visGroupIsSnapTarget): return "IsSnapTarget";
                            case ((short)IVisio.VisCellIndices.visGroupIsTextEditTarget): return "IsTextEditTarget";
                            case ((short)IVisio.VisCellIndices.visGroupSelectMode): return "SelectMode";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowImage):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visImageBlur): return "Blur";
                            case ((short)IVisio.VisCellIndices.visImageBrightness): return "Brightness";
                            case ((short)IVisio.VisCellIndices.visImageContrast): return "Contrast";
                            case ((short)IVisio.VisCellIndices.visImageDenoise): return "Denoise";
                            case ((short)IVisio.VisCellIndices.visImageGamma): return "Gamma";
                            case ((short)IVisio.VisCellIndices.visImageSharpen): return "Sharpen";
                            case ((short)IVisio.VisCellIndices.visImageTransparency): return "Transparency";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowPageLayout):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visPLOAvenueSizeX): return "AvenueSizeX";
                            case ((short)IVisio.VisCellIndices.visPLOAvenueSizeY): return "AvenueSizeY";
                            case ((short)IVisio.VisCellIndices.visPLOBlockSizeX): return "BlockSizeX";
                            case ((short)IVisio.VisCellIndices.visPLOBlockSizeY): return "BlockSizeY";
                            case ((short)IVisio.VisCellIndices.visPLOCtrlAsInput): return "CtrlAsInput";
                            case ((short)IVisio.VisCellIndices.visPLODynamicsOff): return "DynamicsOff";
                            case ((short)IVisio.VisCellIndices.visPLOEnableGrid): return "EnableGrid";
                            case ((short)IVisio.VisCellIndices.visPLOLineAdjustFrom): return "LineAdjustFrom";
                            case ((short)IVisio.VisCellIndices.visPLOLineAdjustTo): return "LineAdjustTo";
                            case ((short)IVisio.VisCellIndices.visPLOJumpCode): return "LineJumpCode";
                            case ((short)IVisio.VisCellIndices.visPLOJumpFactorX): return "LineJumpFactorX";
                            case ((short)IVisio.VisCellIndices.visPLOJumpFactorY): return "LineJumpFactorY";
                            case ((short)IVisio.VisCellIndices.visPLOJumpStyle): return "LineJumpStyle";
                            case ((short)IVisio.VisCellIndices.visPLOLineRouteExt): return "LineRouteExt";
                            case ((short)IVisio.VisCellIndices.visPLOLineToLineX): return "LineToLineX";
                            case ((short)IVisio.VisCellIndices.visPLOLineToLineY): return "LineToLineY";
                            case ((short)IVisio.VisCellIndices.visPLOLineToNodeX): return "LineToNodeX";
                            case ((short)IVisio.VisCellIndices.visPLOLineToNodeY): return "LineToNodeY";
                            case ((short)IVisio.VisCellIndices.visPLOJumpDirX): return "PageLineJumpDirX";
                            case ((short)IVisio.VisCellIndices.visPLOJumpDirY): return "PageLineJumpDirY";
                            case ((short)IVisio.VisCellIndices.visPLOSplit): return "PageShapeSplit";
                            case ((short)IVisio.VisCellIndices.visPLOPlaceDepth): return "PlaceDepth";
                            case ((short)IVisio.VisCellIndices.visPLOPlaceFlip): return "PlaceFlip";
                            case ((short)IVisio.VisCellIndices.visPLOPlaceStyle): return "PlaceStyle";
                            case ((short)IVisio.VisCellIndices.visPLOPlowCode): return "PlowCode";
                            case ((short)IVisio.VisCellIndices.visPLOResizePage): return "ResizePage";
                            case ((short)IVisio.VisCellIndices.visPLORouteStyle): return "RouteStyle";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowPrintProperties):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesLeftMargin): return "PageLeftMargin";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesCenterX): return "CenterX";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesCenterY): return "CenterY";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesOnPage): return "OnPage";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesBottomMargin): return "PageBottomMargin";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesRightMargin): return "PageRightMargin";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPagesX): return "PagesX";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPagesY): return "PagesY";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesTopMargin): return "PageTopMargin";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPaperKind): return "PaperKind";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPrintGrid): return "PrintGrid";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPageOrientation): return "PrintPageOrientation";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesScaleX): return "ScaleX";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesScaleY): return "ScaleY";
                            case ((short)IVisio.VisCellIndices.visPrintPropertiesPaperSource): return "PaperSource";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowPage):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visPageDrawingScale): return "DrawingScale";
                            case ((short)IVisio.VisCellIndices.visPageDrawScaleType): return "DrawingScaleType";
                            case ((short)IVisio.VisCellIndices.visPageDrawSizeType): return "DrawingSizeType";
                            case ((short)IVisio.VisCellIndices.visPageInhibitSnap): return "InhibitSnap";
                            case ((short)IVisio.VisCellIndices.visPageHeight): return "PageHeight";
                            case ((short)IVisio.VisCellIndices.visPageScale): return "PageScale";
                            case ((short)IVisio.VisCellIndices.visPageWidth): return "PageWidth";
                            case ((short)IVisio.VisCellIndices.visPageShdwObliqueAngle): return "ShdwObliqueAngle";
                            case ((short)IVisio.VisCellIndices.visPageShdwOffsetX): return "ShdwOffsetX";
                            case ((short)IVisio.VisCellIndices.visPageShdwOffsetY): return "ShdwOffsetY";
                            case ((short)IVisio.VisCellIndices.visPageShdwScaleFactor): return "ShdwScaleFactor";
                            case ((short)IVisio.VisCellIndices.visPageShdwType): return "ShdwType";
                            case ((short)IVisio.VisCellIndices.visPageUIVisibility): return "UIVisibility";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowLock):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visLockAspect): return "LockAspect";
                            case ((short)IVisio.VisCellIndices.visLockBegin): return "LockBegin";
                            case ((short)IVisio.VisCellIndices.visLockCalcWH): return "LockCalcWH";
                            case ((short)IVisio.VisCellIndices.visLockCrop): return "LockCrop";
                            case ((short)IVisio.VisCellIndices.visLockCustProp): return "LockCustProp";
                            case ((short)IVisio.VisCellIndices.visLockDelete): return "LockDelete";
                            case ((short)IVisio.VisCellIndices.visLockEnd): return "LockEnd";
                            case ((short)IVisio.VisCellIndices.visLockFormat): return "LockFormat";
                            case ((short)IVisio.VisCellIndices.visLockFromGroupFormat): return "LockFromGroupFormat";
                            case ((short)IVisio.VisCellIndices.visLockGroup): return "LockGroup";
                            case ((short)IVisio.VisCellIndices.visLockHeight): return "LockHeight";
                            case ((short)IVisio.VisCellIndices.visLockMoveX): return "LockMoveX";
                            case ((short)IVisio.VisCellIndices.visLockMoveY): return "LockMoveY";
                            case ((short)IVisio.VisCellIndices.visLockRotate): return "LockRotate";
                            case ((short)IVisio.VisCellIndices.visLockSelect): return "LockSelect";
                            case ((short)IVisio.VisCellIndices.visLockTextEdit): return "LockTextEdit";
                            case ((short)IVisio.VisCellIndices.visLockThemeColors): return "LockThemeColors";
                            case ((short)IVisio.VisCellIndices.visLockThemeEffects): return "LockThemeEffects";
                            case ((short)IVisio.VisCellIndices.visLockVtxEdit): return "LockVtxEdit";
                            case ((short)IVisio.VisCellIndices.visLockWidth): return "LockWidth";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowRulerGrid):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visXGridDensity): return "XGridDensity";
                            case ((short)IVisio.VisCellIndices.visXGridOrigin): return "XGridOrigin";
                            case ((short)IVisio.VisCellIndices.visXGridSpacing): return "XGridSpacing";
                            case ((short)IVisio.VisCellIndices.visXRulerDensity): return "XRulerDensity";
                            case ((short)IVisio.VisCellIndices.visXRulerOrigin): return "XRulerOrigin";
                            case ((short)IVisio.VisCellIndices.visYGridDensity): return "YGridDensity";
                            case ((short)IVisio.VisCellIndices.visYGridOrigin): return "YGridOrigin";
                            case ((short)IVisio.VisCellIndices.visYGridSpacing): return "YGridSpacing";
                            case ((short)IVisio.VisCellIndices.visYRulerDensity): return "YRulerDensity";
                            case ((short)IVisio.VisCellIndices.visYRulerOrigin): return "YRulerOrigin";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowTextXForm):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visXFormAngle): return "TxtAngle";
                            case ((short)IVisio.VisCellIndices.visXFormHeight): return "TxtHeight";
                            case ((short)IVisio.VisCellIndices.visXFormLocPinX): return "TxtLocPinX";
                            case ((short)IVisio.VisCellIndices.visXFormLocPinY): return "TxtLocPinY";
                            case ((short)IVisio.VisCellIndices.visXFormPinX): return "TxtPinX";
                            case ((short)IVisio.VisCellIndices.visXFormPinY): return "TxtPinY";
                            case ((short)IVisio.VisCellIndices.visXFormWidth): return "TxtWidth";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowText):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visTxtBlkBottomMargin): return "BottomMargin";
                            case ((short)IVisio.VisCellIndices.visTxtBlkDefaultTabStop): return "DefaultTabStop";
                            case ((short)IVisio.VisCellIndices.visTxtBlkLeftMargin): return "LeftMargin";
                            case ((short)IVisio.VisCellIndices.visTxtBlkRightMargin): return "RightMargin";
                            case ((short)IVisio.VisCellIndices.visTxtBlkBkgnd): return "TextBkgnd";
                            case ((short)IVisio.VisCellIndices.visTxtBlkBkgndTrans): return "TextBkgndTrans";
                            case ((short)IVisio.VisCellIndices.visTxtBlkDirection): return "TextDirection";
                            case ((short)IVisio.VisCellIndices.visTxtBlkTopMargin): return "TopMargin";
                            case ((short)IVisio.VisCellIndices.visTxtBlkVerticalAlign): return "VerticalAlign";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowStyle):
                    {
                        switch (src.Cell)
                        {

                            case ((short)IVisio.VisCellIndices.visStyleIncludesFill): return "EnableFillProps";
                            case ((short)IVisio.VisCellIndices.visStyleIncludesLine): return "EnableLineProps";
                            case ((short)IVisio.VisCellIndices.visStyleIncludesText): return "EnableTextProps";
                            case ((short)IVisio.VisCellIndices.visStyleHidden): return "HideForApply";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowXForm1D):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.vis1DBeginX): return "BeginX";
                            case ((short)IVisio.VisCellIndices.vis1DBeginY): return "BeginY";
                            case ((short)IVisio.VisCellIndices.vis1DEndX): return "EndX";
                            case ((short)IVisio.VisCellIndices.vis1DEndY): return "EndY";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowShapeLayout):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visSLOConFixedCode): return "ConFixedCode";
                            case ((short)IVisio.VisCellIndices.visSLOJumpCode): return "ConLineJumpCode";
                            case ((short)IVisio.VisCellIndices.visSLOJumpDirX): return "ConLineJumpDirX";
                            case ((short)IVisio.VisCellIndices.visSLOJumpDirY): return "ConLineJumpDirY";
                            case ((short)IVisio.VisCellIndices.visSLOJumpStyle): return "ConLineJumpStyle";
                            case ((short)IVisio.VisCellIndices.visSLOLineRouteExt): return "ConLineRouteExt";
                            case ((short)IVisio.VisCellIndices.visSLOFixedCode): return "ShapeFixedCode";
                            case ((short)IVisio.VisCellIndices.visSLOPermeablePlace): return "ShapePermeablePlace";
                            case ((short)IVisio.VisCellIndices.visSLOPermX): return "ShapePermeableX";
                            case ((short)IVisio.VisCellIndices.visSLOPermY): return "ShapePermeableY";
                            case ((short)IVisio.VisCellIndices.visSLOPlaceFlip): return "ShapePlaceFlip";
                            case ((short)IVisio.VisCellIndices.visSLOPlaceStyle): return "ShapePlaceStyle";
                            case ((short)IVisio.VisCellIndices.visSLOPlowCode): return "ShapePlowCode";
                            case ((short)IVisio.VisCellIndices.visSLORouteStyle): return "ShapeRouteStyle";
                            case ((short)IVisio.VisCellIndices.visSLOSplit): return "ShapeSplit";
                            case ((short)IVisio.VisCellIndices.visSLOSplittable): return "ShapeSplittable";
                            default:
                                break;
                        }
                        break;
                    }
                default:
                    break;
            }
            return null;

        }

        public static SRC? TryGetSRCFromName(string name)
        {
            var dic = NameToSRCDictionary;
            SRC src;
            bool found_value = dic.TryGetValue(name, out src);

            if (found_value)
            {
                return src;
            }

            int dot_pos = name.IndexOf('.');
            if (dot_pos >= 0)
            {
                string left_of_dot = name.Substring(0, dot_pos);

                int left_bracket_pos = name.IndexOf('[');
                if (left_bracket_pos >= 0)
                {
                    string unbracketed_name = name.Substring(0, left_bracket_pos);
                    int len = left_bracket_pos - dot_pos - 1;
                    string between = name.Substring(dot_pos+1, len);

                    int right_bracket_pos = name.IndexOf(']');
                    if (right_bracket_pos > 0)
                    {
                        int between_brackets_len = right_bracket_pos - left_bracket_pos - 1;
                        string between_brackets_str = name.Substring(left_bracket_pos + 1, between_brackets_len);

                        if ((left_of_dot == "Char") || (left_of_dot == "Para"))
                        {
                            var x = TryGetSRCFromName(unbracketed_name);
                            if (x.HasValue)
                            {
                                int bracket_int = int.Parse(between_brackets_str);
                                var y = x.Value.ForRow((short) (bracket_int -1) );
                                return y;

                            }
                        }

                    }




                }

                if (left_of_dot == "Para")
                {
                }
            }

            return null;
        }

        public static SRC GetSRCFromName(string name)
        {
            var src = TryGetSRCFromName(name);
            if (src.HasValue)
            {
                return src.Value;
            }

            string msg = string.Format("Cannot identify indices for cell with name \"{0}\"", name);
            throw new AutomationException(msg);
        }

        private static Dictionary<string, SRC> NameToSRCDictionary
        {
            get
            {
                if (simple_name_to_src_map == null)
                {
                    CreateNameToSRCDictionary();
                }
                return simple_name_to_src_map;
            }
        }

        private static void CreateNameToSRCDictionary()
        {
            simple_name_to_src_map = new Dictionary<string, SRC>(StringComparer.OrdinalIgnoreCase)
                                             {
                                                 {"PinX",SRCConstants.PinX},
                                                 {"PinY",SRCConstants.PinY},
                                                 {"LocPinX",SRCConstants.LocPinX},
                                                 {"LocPinY",SRCConstants.LocPinY},
                                                 {"Width",SRCConstants.Width},
                                                 {"Height",SRCConstants.Height},
                                                 {"Angle",SRCConstants.Angle},
                                                 {"FlipX",SRCConstants.FlipX},
                                                 {"FlipY",SRCConstants.FlipY},
                                                 {"ResizeMode",SRCConstants.ResizeMode},


                                                 {"FillBkgnd",SRCConstants.FillBkgnd},
                                                 {"FillBkgndTrans",SRCConstants.FillBkgndTrans},
                                                 {"FillForegnd",SRCConstants.FillForegnd},
                                                 {"FillForegndTrans",SRCConstants.FillForegndTrans},
                                                 {"FillPattern",SRCConstants.FillPattern},
                                                 {"ShapeShdwObliqueAngle",SRCConstants.ShapeShdwObliqueAngle},
                                                 {"ShapeShdwOffsetX",SRCConstants.ShapeShdwOffsetX},
                                                 {"ShapeShdwOffsetY",SRCConstants.ShapeShdwOffsetY},
                                                 {"ShapeShdwScaleFactor",SRCConstants.ShapeShdwScaleFactor},
                                                 {"ShapeShdwType",SRCConstants.ShapeShdwType},
                                                 {"ShdwBkgnd",SRCConstants.ShdwBkgnd},
                                                 {"ShdwBkgndTrans",SRCConstants.ShdwBkgndTrans},
                                                 {"ShdwForegnd",SRCConstants.ShdwForegnd},
                                                 {"ShdwForegndTrans",SRCConstants.ShdwForegndTrans},
                                                 {"ShdwPattern",SRCConstants.ShdwPattern},

                                                 {"LineCap",SRCConstants.LineCap},
                                                 {"LineColor",SRCConstants.LineColor},
                                                 {"LineColorTrans",SRCConstants.LineColorTrans},
                                                 {"LineWeight",SRCConstants.LineWeight},
                                                 {"LinePattern",SRCConstants.LinePattern},
                                                 {"Rounding",SRCConstants.Rounding},
                                                 {"BeginArrow",SRCConstants.BeginArrow},
                                                 {"BeginArrowSize",SRCConstants.BeginArrowSize},
                                                 {"EndArrow",SRCConstants.EndArrow},
                                                 {"EndArrowSize",SRCConstants.EndArrowSize},

                                                 {"BeginX",SRCConstants.BeginX},
                                                 {"BeginY",SRCConstants.BeginY},
                                                 {"EndX",SRCConstants.EndX},
                                                 {"EndY",SRCConstants.EndY},
                                                 
                                                 {"Char.Case",SRCConstants.Char_Case},
                                                 {"Char.Color",SRCConstants.Char_Color},
                                                 {"Char.ColorTrans",SRCConstants.Char_ColorTrans},
                                                 {"Char.DblUnderline",SRCConstants.Char_DblUnderline},
                                                 {"Char.DoubleStrikethrough",SRCConstants.Char_DoubleStrikethrough},
                                                 {"Char.Font",SRCConstants.Char_Font},
                                                 {"Char.FontScale",SRCConstants.Char_FontScale},
                                                 {"Char.Letterspace",SRCConstants.Char_Letterspace},
                                                 {"Char.Overline",SRCConstants.Char_Overline},
                                                 {"Char.Size",SRCConstants.Char_Size},
                                                 {"Char.Strikethru",SRCConstants.Char_Strikethru},
                                                 {"Char.Style",SRCConstants.Char_Style},
                                                 {"Char.Pos",SRCConstants.Char_Pos},
                                                 {"Char.RTLText",SRCConstants.RTLText},
                                                 {"Char.UseVertical",SRCConstants.UseVertical},

                                                 //glueinfo
                                                 {"BegTrigger",SRCConstants.BegTrigger},
                                                 {"EndTrigger",SRCConstants.EndTrigger},
                                                 {"GlueType",SRCConstants.GlueType},
                                                 {"WalkPreference",SRCConstants.WalkPreference},

                                                 // group
                                                 {"DisplayMode",SRCConstants.DisplayMode},
                                                 {"DontMoveChildren",SRCConstants.DontMoveChildren},
                                                 {"IsDropTarget",SRCConstants.IsDropTarget},
                                                 {"IsSnapTarget",SRCConstants.IsSnapTarget},
                                                 {"IsTextEditTarget",SRCConstants.IsTextEditTarget},
                                                 {"SelectMode",SRCConstants.SelectMode},

                                                 // misc
                                                 {"Calendar",SRCConstants.Calendar},
                                                 {"Comment",SRCConstants.Comment},
                                                 {"DropOnPageScale",SRCConstants.DropOnPageScale},
                                                 {"DynFeedback",SRCConstants.DynFeedback},
                                                 {"HideText",SRCConstants.HideText},
                                                 {"IsDropSource",SRCConstants.IsDropSource},
                                                 {"LangID",SRCConstants.LangID},
                                                 {"LocalizeMerge",SRCConstants.LocalizeMerge},
                                                 {"NoAlignBox",SRCConstants.NoAlignBox},
                                                 {"NoCtlHandles",SRCConstants.NoCtlHandles},
                                                 {"NoLiveDynamics",SRCConstants.NoLiveDynamics},
                                                 {"NonPrinting",SRCConstants.NonPrinting},
                                                 {"NoObjHandles",SRCConstants.NoObjHandles},
                                                 {"ObjType",SRCConstants.ObjType},
                                                 {"UpdateAlignBox",SRCConstants.UpdateAlignBox},




                                                 {"Para.Bullet",SRCConstants.Para_Bullet},
                                                 {"Para.BulletFont",SRCConstants.Para_BulletFont},
                                                 {"Para.BulletFontSize",SRCConstants.Para_BulletFontSize},
                                                 {"Para.BulletStr",SRCConstants.Para_BulletStr},
                                                 {"Para.Flags",SRCConstants.Para_Flags},
                                                 {"Para.HorzAlign",SRCConstants.Para_HorzAlign},
                                                 {"Para.IndFirst",SRCConstants.Para_IndFirst},
                                                 {"Para.IndLeft",SRCConstants.Para_IndLeft},
                                                 {"Para.IndRight",SRCConstants.Para_IndRight},
                                                 {"Para.LocBulletFont",SRCConstants.Para_LocalizeBulletFont},
                                                 {"Para.SpAfter",SRCConstants.Para_SpAfter},
                                                 {"Para.SpBefore",SRCConstants.Para_SpBefore},
                                                 {"Para.SpLine",SRCConstants.Para_SpLine},
                                                 {"Para.TextPosAfterBullet",SRCConstants.Para_TextPosAfterBullet},
                                                                      
                                                 {"LockAspect",SRCConstants.LockAspect},
                                                 {"LockBegin",SRCConstants.LockBegin},
                                                 {"LockCalcWH",SRCConstants.LockCalcWH},
                                                 {"LockCrop",SRCConstants.LockCrop},
                                                 {"LockCustProp",SRCConstants.LockCustProp},
                                                 {"LockDelete",SRCConstants.LockDelete},
                                                 {"LockEnd",SRCConstants.LockEnd},
                                                 {"LockFormat",SRCConstants.LockFormat},
                                                 {"LockFromGroupFormat",SRCConstants.LockFromGroupFormat},
                                                 {"LockGroup",SRCConstants.LockGroup},
                                                 {"LockHeight",SRCConstants.LockHeight},
                                                 {"LockMoveX",SRCConstants.LockMoveX},
                                                 {"LockMoveY",SRCConstants.LockMoveY},
                                                 {"LockRotate",SRCConstants.LockRotate},
                                                 {"LockSelect",SRCConstants.LockSelect},
                                                 {"LockTextEdit",SRCConstants.LockTextEdit},
                                                 {"LockThemeColors",SRCConstants.LockThemeColors},
                                                 {"LockThemeEffects",SRCConstants.LockThemeEffects},
                                                 {"LockVtxEdit",SRCConstants.LockVtxEdit},
                                                 {"LockWidth",SRCConstants.LockWidth},
                                                                      
                                                 {"TxtAngle",SRCConstants.TxtAngle },
                                                 {"TxtHeight",SRCConstants.TxtHeight },
                                                 {"TxtLocPinX",SRCConstants.TxtLocPinX},
                                                 {"TxtLocPinY",SRCConstants.TxtLocPinY},
                                                 {"TxtPinX",SRCConstants.TxtPinX },
                                                 {"TxtPinY",SRCConstants.TxtPinY  },
                                                 {"TxtWidth",SRCConstants.TxtWidth },
                                                                      
                                                 {"BottomMargin",SRCConstants.BottomMargin },
                                                 {"DefaultTabstop",SRCConstants.DefaultTabstop},
                                                 {"LeftMargin",SRCConstants.LeftMargin },
                                                 {"RightMargin",SRCConstants.RightMargin  },
                                                 {"TextBkgnd",SRCConstants.TextBkgnd },
                                                 {"TextBkgndTrans",SRCConstants.TextBkgndTrans},
                                                 {"TextDirection",SRCConstants.TextDirection },
                                                 {"TopMargin",SRCConstants.TopMargin },
                                                 {"VerticalAlign",SRCConstants.VerticalAlign },
                                                                      
                                                 {"ConFixedCode",SRCConstants.ConFixedCode},
                                                 {"ConLineJumpCode",SRCConstants.ConLineJumpCode},
                                                 {"ConLineJumpDirX",SRCConstants.ConLineJumpDirX},
                                                 {"ConLineJumpDirY",SRCConstants.ConLineJumpDirY},
                                                 {"ConLineJumpStyle",SRCConstants.ConLineJumpStyle},
                                                 {"ConLineRouteExt",SRCConstants.ConLineRouteExt},
                                                 {"ShapeFixedCode",SRCConstants.ShapeFixedCode},
                                                 {"ShapePermeablePlace",SRCConstants.ShapePermeablePlace},
                                                 {"ShapePermeableX",SRCConstants.ShapePermeableX},
                                                 {"ShapePermeableY",SRCConstants.ShapePermeableY},
                                                 {"ShapePlaceFlip",SRCConstants.ShapePlaceFlip},
                                                 {"ShapePlaceStyle",SRCConstants.ShapePlaceStyle},
                                                 {"ShapePlowCode",SRCConstants.ShapePlowCode},
                                                 {"ShapeRouteStyle",SRCConstants.ShapeRouteStyle},
                                                 {"ShapeSplit",SRCConstants.ShapeSplit},
                                                 {"ShapeSplittable",SRCConstants.ShapeSplittable},
                                             };
            
        }

        private static Dictionary<string, SRC> simple_name_to_src_map;
    }
}