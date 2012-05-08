import win32com.client 

class SRC(object) :

    def __init__( self, s, r, c ) :
        self.Section = s
        self.Row = r
        self.Cell = c

    def __str__(self) :
        return "SRC(%s,%s,%s)" % (self.Section,self.Row,self.Cell)

class SIDSRC(object) :

    def __init__( self, id, src ) :
        self.ShapeID = id
        self.SRC = src

    def __str__(self) :
        return "SIDSRC(%s,%s,%s,%s)" % (self.ShapeID,self.SRC.Section,self.SRC.Row,self.SRC.Cell)

class ShapeSheetUtil(object) :

    @staticmethod
    def BuildSIDSRCStream( sidsrcs ) :
        stream = []
        for sidsrc in sidsrcs :
            stream.append(sidsrc.ShapeID)
            stream.append(sidsrc.SRC.Section)
            stream.append(sidsrc.SRC.Row)
            stream.append(sidsrc.SRC.Cell)
        return stream

class Query(object) :

    # page.GetFormulas: http://msdn.microsoft.com/en-us/library/ff768473.aspx
    # page.GetResults: http://msdn.microsoft.com/en-us/library/ff766481.aspx

    def __init__(self) :
        self.items = []
        self.GetResultsFlags = 0
        self.UnitCodes = None

    def Add(self, id, src) :
        sidsrc = SIDSRC(id,src)
        self.items.append(sidsrc)

    def GetFormulas(self, page) :
        formulas,results= self.__getdata(page,False,True)
        return formulas

    def GetResults(self, page) :
        formulas,results= self.__getdata(page,False,True)
        return results

    def GetFormulasAndResults(self, page) :
        formulas,results= self.__getdata(page,True,True)
        return formulas,results

    def __getdata(self, page, getformulas, getresults) :
        stream = ShapeSheetUtil.BuildSIDSRCStream( self.items )
        formulas = None
        results = None

        if (getformulas) : 
            formulas = page.GetFormulas(stream)
            assert( len(formulas) == len(self.items))
        if (getresults) : 
            results = page.GetResults(stream, self.GetResultsFlags, self.UnitCodes)
            assert( len(results) == len(self.items))
        
        return (formulas,results)

    @staticmethod
    def __tabulate( numcols, data) :
        container = []
        curlist = None
        for i,d in enumerate(data) :
            if (i%numcols==0) : 
                cur_list = []
            cur_list.append(d)
            if (i%numcols==numcols-1) :
                container.append(cur_list)

        total_count = sum( len(r) for r in container)
        assert(total_count==len(data))
        return container

    @staticmethod
    def __buildquery_internal(shapeids,srcs) :
        q = Query()
        for shapeid in shapeids:
            for src in srcs:
                q.Add(shapeid,src)
        return q

    @staticmethod
    def QueryFormulas(page,shapeids,srcs) :
        q = Query.__buildquery_internal(shapeids,srcs)
        formulas = q.GetFormulas(page)
        return Query.__tabulate(len(srcs),formulas)

    @staticmethod
    def QueryResults(page,shapeids,srcs) :
        q = Query.__buildquery_internal(shapeids,srcs)
        formulas = q.GetResults(page)
        return Query.__tabulate(len(srcs),formulas)

class Update(object) :

    def __init__(self) :
        self.items = []
        self.SetFormulasFlags = 0

    def Add(self, id, src, formula ) :
        item = (SIDSRC(id,src),formula)
        self.items.append(item)

    def SetFormulas(self, page) :
        if (len(self.items)<1) :
            return (0, [])
        stream = ShapeSheetUtil.BuildSIDSRCStream( (s[0] for s in self.items) )
        formulas = []
        for (sidsrc,formula) in self.items :
            formulas.append(formula)
        result = page.SetFormulas(stream, formulas, self.SetFormulasFlags)
        return result

class SRCConstants(object):

    __SEC__ = win32com.client.constants
    __ROW__ = win32com.client.constants
    __CEL__ = win32com.client.constants

    #  Actions
    Actions_Action  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionAction)
    Actions_BeginGroup  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionBeginGroup)
    Actions_ButtonFace  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionButtonFace)
    Actions_Checked  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionChecked)
    Actions_Disabled  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionDisabled)
    Actions_Invisible  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionInvisible)
    Actions_Menu  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionMenu)
    Actions_ReadOnly  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionReadOnly)
    Actions_SortKey  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionSortKey)
    Actions_TagName  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionTagName)
    Actions_FlyoutChild  = SRC(__SEC__.visSectionAction, __ROW__.visRowAction, __CEL__.visActionFlyoutChild) #  new for visio 2010

    #  Alignment
    AlignBottom  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignBottom)
    AlignCenter  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignCenter)
    AlignLeft  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignLeft)
    AlignMiddle  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignMiddle)
    AlignRight  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignRight)
    AlignTop  = SRC(__SEC__.visSectionObject, __ROW__.visRowAlign, __CEL__.visAlignTop)

    #  Annotation
    Annotation_Comment  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationComment)
    Annotation_Date  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationDate)
    Annotation_LangID  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationLangID)
    Annotation_MarkerIndex  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationMarkerIndex)
    Annotation_X  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationX)
    Annotation_Y  = SRC(__SEC__.visSectionAnnotation, __ROW__.visRowAnnotation, __CEL__.visAnnotationY)

    #  Character
    Char_AsianFont  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterAsianFont)
    Char_Case  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterCase)
    Char_Color  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterColor)
    Char_ComplexScriptFont  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterComplexScriptFont)
    Char_ComplexScriptSize  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterComplexScriptSize)
    Char_DoubleStrikethrough  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterDoubleStrikethrough)
    Char_DblUnderline  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterDblUnderline)
    Char_Font  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterFont)
    Char_LangID  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterLangID)
    Char_Locale  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterLocale)
    Char_LocalizeFont  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterLocalizeFont)
    Char_Overline  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterOverline)
    Char_Perpendicular  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterPerpendicular)
    Char_Pos  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterPos)
    Char_RTLText  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterRTLText)
    Char_FontScale  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterFontScale)
    Char_Size  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterSize)
    Char_Letterspace  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterLetterspace)
    Char_Strikethru  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterStrikethru)
    Char_Style  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterStyle)
    Char_ColorTrans  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterColorTrans)
    Char_UseVertical  = SRC(__SEC__.visSectionCharacter, __ROW__.visRowCharacter, __CEL__.visCharacterUseVertical)

    #  Connections
    Connections_D  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visCnnctD)
    Connections_DirX  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visCnnctDirX)
    Connections_DirY  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visCnnctDirY)
    Connections_Type  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visCnnctType)
    Connections_X  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visX)
    Connections_Y  = SRC(__SEC__.visSectionConnectionPts, __ROW__.visRowConnectionPts, __CEL__.visY)

    #  Controls
    Controls_CanGlue  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlGlue)
    Controls_Tip  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlTip)
    Controls_XCon  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlXCon)
    Controls_X  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlX)
    Controls_XDyn  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlXDyn)
    Controls_YCon  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlYCon)
    Controls_Y  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlY)
    Controls_YDyn  = SRC(__SEC__.visSectionControls, __ROW__.visRowControl, __CEL__.visCtlYDyn)

    #  Document Properties
    AddMarkup  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocAddMarkup)
    DocLangID  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocLangID)
    LockPreview  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocLockPreview)
    OutputFormat  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocOutputFormat)
    PreviewQuality  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocPreviewQuality)
    PreviewScope  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocPreviewScope)
    ViewMarkup  = SRC(__SEC__.visSectionObject, __ROW__.visRowDoc, __CEL__.visDocViewMarkup)

    #  Events
    EventDblClick  = SRC(__SEC__.visSectionObject, __ROW__.visRowEvent, __CEL__.visEvtCellDblClick)
    EventDrop  = SRC(__SEC__.visSectionObject, __ROW__.visRowEvent, __CEL__.visEvtCellDrop)
    EventMultiDrop  = SRC(__SEC__.visSectionObject, __ROW__.visRowEvent, __CEL__.visEvtCellMultiDrop)
    EventXFMod  = SRC(__SEC__.visSectionObject, __ROW__.visRowEvent, __CEL__.visEvtCellXFMod)
    TheText  = SRC(__SEC__.visSectionObject, __ROW__.visRowEvent, __CEL__.visEvtCellTheText)

    #  ForeignImageInfo
    ImgHeight  = SRC(__SEC__.visSectionObject, __ROW__.visRowForeign, __CEL__.visFrgnImgHeight)
    ImgOffsetX  = SRC(__SEC__.visSectionObject, __ROW__.visRowForeign, __CEL__.visFrgnImgOffsetX)
    ImgOffsetY  = SRC(__SEC__.visSectionObject, __ROW__.visRowForeign, __CEL__.visFrgnImgOffsetY)
    ImgWidth  = SRC(__SEC__.visSectionObject, __ROW__.visRowForeign, __CEL__.visFrgnImgWidth)

    #  Geometry
    Geometry_A  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visBow)
    Geometry_B  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visControlX)
    Geometry_C  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visEccentricityAngle)
    Geometry_D  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visAspectRatio)
    Geometry_E  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visNURBSData)
    Geometry_X  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visX)
    Geometry_Y  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowVertex, __CEL__.visY)
    Geometry_NoFill  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowComponent, __CEL__.visCompNoFill)
    Geometry_NoLine  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowComponent, __CEL__.visCompNoLine)
    Geometry_NoShow  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowComponent, __CEL__.visCompNoShow)
    Geometry_NoSnap  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowComponent, __CEL__.visCompNoSnap)
    Geometry_NoQuickDrag  = SRC(__SEC__.visSectionFirstComponent, __ROW__.visRowComponent, __CEL__.visCompNoQuickDrag)

    #  Fill Format
    FillBkgnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillBkgnd)
    FillBkgndTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillBkgndTrans)
    FillForegnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillForegnd)
    FillForegndTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillForegndTrans)
    FillPattern  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillPattern)
    ShapeShdwObliqueAngle  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwObliqueAngle)
    ShapeShdwOffsetX  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwOffsetX)
    ShapeShdwOffsetY  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwOffsetY)
    ShapeShdwScaleFactor  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwScaleFactor)
    ShapeShdwType  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwType)
    ShdwBkgnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwBkgnd)
    ShdwBkgndTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwBkgndTrans)
    ShdwForegnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwForegnd)
    ShdwForegndTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwForegndTrans)
    ShdwPattern  = SRC(__SEC__.visSectionObject, __ROW__.visRowFill, __CEL__.visFillShdwPattern)

    #  GlueInfo
    BegTrigger  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visBegTrigger)
    EndTrigger  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visEndTrigger)
    GlueType  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visGlueType)
    WalkPreference  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visWalkPref)

    #  GroupProperties
    DisplayMode  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupDisplayMode)
    DontMoveChildren  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupDontMoveChildren)
    IsDropTarget  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupIsDropTarget)
    IsSnapTarget  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupIsSnapTarget)
    IsTextEditTarget  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupIsTextEditTarget)
    SelectMode  = SRC(__SEC__.visSectionObject, __ROW__.visRowGroup, __CEL__.visGroupSelectMode)

    #  Hyperlinks
    Hyperlink_Address  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkAddress)
    Hyperlink_Default  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkDefault)
    Hyperlink_Description  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkDescription)
    Hyperlink_ExtraInfo  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkExtraInfo)
    Hyperlink_Frame  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkFrame)
    Hyperlink_Invisible  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkInvisible)
    Hyperlink_NewWindow  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkNewWin)
    Hyperlink_SortKey  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkSortKey)
    Hyperlink_SubAddress  = SRC(__SEC__.visSectionHyperlink, __ROW__.visRow1stHyperlink, __CEL__.visHLinkSubAddress)

    #  Image Properties
    Blur  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageBlur)
    Brightness  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageBrightness)
    Contrast  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageContrast)
    Denoise  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageDenoise)
    Gamma  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageGamma)
    Sharpen  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageSharpen)
    Transparency  = SRC(__SEC__.visSectionObject, __ROW__.visRowImage, __CEL__.visImageTransparency)

    #  Line format
    BeginArrow  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineBeginArrow)
    BeginArrowSize  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineBeginArrowSize)
    EndArrow  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineEndArrow)
    EndArrowSize  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineEndArrowSize)
    LineCap  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineEndCap)
    LineColor  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineColor)
    LineColorTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineColorTrans)
    LinePattern  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLinePattern)
    LineWeight  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineWeight)
    Rounding  = SRC(__SEC__.visSectionObject, __ROW__.visRowLine, __CEL__.visLineRounding)

    #  Miscellaneous
    Calendar  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visObjCalendar)
    Comment  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visComment)
    DropOnPageScale  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visObjDropOnPageScale)
    DynFeedback  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visDynFeedback)
    IsDropSource  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visDropSource)
    LangID  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visObjLangID)
    LocalizeMerge  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visObjLocalizeMerge)
    NoAlignBox  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visNoAlignBox)
    NoCtlHandles  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visNoCtlHandles)
    NoLiveDynamics  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visNoLiveDynamics)
    NonPrinting  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visNonPrinting)
    NoObjHandles  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visNoObjHandles)
    ObjType  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visLOFlags)
    UpdateAlignBox  = SRC(__SEC__.visSectionObject, __ROW__.visRowMisc, __CEL__.visUpdateAlignBox)

    #  1d endpoints
    BeginX  = SRC(__SEC__.visSectionObject, __ROW__.visRowXForm1D, __CEL__.vis1DBeginX)
    BeginY  = SRC(__SEC__.visSectionObject, __ROW__.visRowXForm1D, __CEL__.vis1DBeginY)
    EndX  = SRC(__SEC__.visSectionObject, __ROW__.visRowXForm1D, __CEL__.vis1DEndX)
    EndY  = SRC(__SEC__.visSectionObject, __ROW__.visRowXForm1D, __CEL__.vis1DEndY)

    #  page layout
    AvenueSizeX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOAvenueSizeX)
    AvenueSizeY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOAvenueSizeY)
    BlockSizeX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOBlockSizeX)
    BlockSizeY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOBlockSizeY)
    CtrlAsInput  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOCtrlAsInput)
    DynamicsOff  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLODynamicsOff)
    EnableGrid  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOEnableGrid)
    LineAdjustFrom  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineAdjustFrom)
    LineAdjustTo  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineAdjustTo)
    LineJumpCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpCode)
    LineJumpFactorX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpFactorX)
    LineJumpFactorY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpFactorY)
    LineJumpStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpStyle)
    LineRouteExt  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineRouteExt)
    LineToLineX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineToLineX)
    LineToLineY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineToLineY)
    LineToNodeX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineToNodeX)
    LineToNodeY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOLineToNodeY)
    PageLineJumpDirX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpDirX)
    PageLineJumpDirY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOJumpDirY)
    PageShapeSplit  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOSplit)
    PlaceDepth  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOPlaceDepth)
    PlaceFlip  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOPlaceFlip)
    PlaceStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOPlaceStyle)
    PlowCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOPlowCode)
    ResizePage  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOResizePage)
    RouteStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLORouteStyle)

    AvoidPageBreaks  = SRC(__SEC__.visSectionObject, __ROW__.visRowPageLayout, __CEL__.visPLOAvoidPageBreaks) #  new in Visio 2010

    #  print properties
    PageLeftMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesLeftMargin)
    CenterX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesCenterX)
    CenterY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesCenterY)
    OnPage  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesOnPage)
    PageBottomMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesBottomMargin)
    PageRightMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesRightMargin)
    PagesX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPagesX)
    PagesY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPagesY)
    PageTopMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesTopMargin)
    PaperKind  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPaperKind)
    PrintGrid  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPrintGrid)
    PrintPageOrientation  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPageOrientation)
    ScaleX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesScaleX)
    ScaleY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesScaleY)
    PaperSource  = SRC(__SEC__.visSectionObject, __ROW__.visRowPrintProperties, __CEL__.visPrintPropertiesPaperSource)

    #  page properties
    DrawingScale  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageDrawingScale)
    DrawingScaleType  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageDrawScaleType)
    DrawingSizeType  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageDrawSizeType)
    InhibitSnap  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageInhibitSnap)
    PageHeight  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageHeight)
    PageScale  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageScale)
    PageWidth  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageWidth)
    ShdwObliqueAngle  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageShdwObliqueAngle)
    ShdwOffsetX  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageShdwOffsetX)
    ShdwOffsetY  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageShdwOffsetY)
    ShdwScaleFactor  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageShdwScaleFactor)
    ShdwType  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageShdwType)
    UIVisibility  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageUIVisibility)
    DrawingResizeType  = SRC(__SEC__.visSectionObject, __ROW__.visRowPage, __CEL__.visPageDrawResizeType) #  new in Visio 2010

    #  paragraph
    Para_Bullet  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visBulletIndex)
    Para_BulletFont  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visBulletFont)
    Para_BulletFontSize  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visBulletFontSize)
    Para_BulletStr  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visBulletString)
    Para_Flags  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visFlags)
    Para_HorzAlign  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visHorzAlign)
    Para_IndFirst  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visIndentFirst)
    Para_IndLeft  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visIndentLeft)
    Para_IndRight  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visIndentRight)
    Para_LocalizeBulletFont  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visLocalizeBulletFont)
    Para_SpAfter  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visSpaceAfter)
    Para_SpBefore  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visSpaceBefore)
    Para_SpLine  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visSpaceLine)
    Para_TextPosAfterBullet  = SRC(__SEC__.visSectionParagraph, __ROW__.visRowParagraph, __CEL__.visTextPosAfterBullet)

    #  protection
    LockAspect  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockAspect)
    LockBegin  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockBegin)
    LockCalcWH  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockCalcWH)
    LockCrop  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockCrop)
    LockCustProp  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockCustProp)
    LockDelete  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockDelete)
    LockEnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockEnd)
    LockFormat  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockFormat)
    LockFromGroupFormat  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockFromGroupFormat)
    LockGroup  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockGroup)
    LockHeight  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockHeight)
    LockMoveX  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockMoveX)
    LockMoveY  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockMoveY)
    LockRotate  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockRotate)
    LockSelect  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockSelect)
    LockTextEdit  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockTextEdit)
    LockThemeColors  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockThemeColors)
    LockThemeEffects  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockThemeEffects)
    LockVtxEdit  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockVtxEdit)
    LockWidth  = SRC(__SEC__.visSectionObject, __ROW__.visRowLock, __CEL__.visLockWidth)

    #  ruler and grid
    XGridDensity  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visXGridDensity)
    XGridOrigin  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visXGridOrigin)
    XGridSpacing  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visXGridSpacing)
    XRulerDensity  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visXRulerDensity)
    XRulerOrigin  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visXRulerOrigin)
    YGridDensity  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visYGridDensity)
    YGridOrigin  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visYGridOrigin)
    YGridSpacing  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visYGridSpacing)
    YRulerDensity  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visYRulerDensity)
    YRulerOrigin  = SRC(__SEC__.visSectionObject, __ROW__.visRowRulerGrid, __CEL__.visYRulerOrigin)

    #  Shape Tranform
    Angle  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormAngle)
    FlipX  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormFlipX)
    FlipY  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormFlipY)
    Height  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormHeight)
    LocPinX  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormLocPinX)
    LocPinY  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormLocPinY)
    PinX  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormPinX)
    PinY  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormPinY)
    ResizeMode  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormResizeMode)
    Width  = SRC(__SEC__.visSectionObject, __ROW__.visRowXFormOut, __CEL__.visXFormWidth)

    #  reviewer
    Reviewer_Color  = SRC(__SEC__.visSectionReviewer, __ROW__.visRowReviewer, __CEL__.visReviewerColor)
    Reviewer_Initials  = SRC(__SEC__.visSectionReviewer, __ROW__.visRowReviewer, __CEL__.visReviewerInitials)
    Reviewer_Name  = SRC(__SEC__.visSectionReviewer, __ROW__.visRowReviewer, __CEL__.visReviewerName)

    #  shape data
    Prop_SortKey  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsSortKey)
    Prop_Ask  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsAsk)
    Prop_Calendar  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsCalendar)
    Prop_Format  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsFormat)
    Prop_Invisible  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsInvis)
    Prop_Label  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsLabel)
    Prop_LangID  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsLangID)
    Prop_Prompt  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsPrompt)
    Prop_Type  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsType)
    Prop_Value  = SRC(__SEC__.visSectionProp, __ROW__.visRowProp, __CEL__.visCustPropsValue)

    #  Layers
    Layers_Active  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerActive)
    Layers_Color  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerColor)
    Layers_Glue  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerGlue)
    Layers_Locked  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerLock)
    Layers_Print  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visDocPreviewScope)
    Layers_Snap  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerSnap)
    Layers_ColorTrans  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerColorTrans)
    Layers_Visible  = SRC(__SEC__.visSectionLayer, __ROW__.visRowLayer, __CEL__.visLayerVisible)

    # text transform
    TxtAngle  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormAngle)
    TxtHeight  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormHeight)
    TxtLocPinX  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormLocPinX)
    TxtLocPinY  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormLocPinY)
    TxtPinX  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormPinX)
    TxtPinY  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormPinY)
    TxtWidth  = SRC(__SEC__.visSectionObject, __ROW__.visRowTextXForm, __CEL__.visXFormWidth)

    #  user defined cells
    User_Prompt  = SRC(__SEC__.visSectionUser, __ROW__.visRowUser, __CEL__.visUserPrompt)
    User_Value  = SRC(__SEC__.visSectionUser, __ROW__.visRowUser, __CEL__.visUserValue)

    #  Fields
    Fields_Calendar  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldCalendar)
    Fields_Format  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldFormat)
    Fields_ObjectKind  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldObjectKind)
    Fields_Type  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldType)
    Fields_UICat  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldUICategory)
    Fields_UICod  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldUICode)
    Fields_UIFmt  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldUIFormat)
    Fields_Value  = SRC(__SEC__.visSectionTextField, __ROW__.visRowField, __CEL__.visFieldCell)

    #  text block format
    BottomMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkBottomMargin)
    DefaultTabStop  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkDefaultTabStop)
    LeftMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkLeftMargin)
    RightMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkRightMargin)
    TextBkgnd  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkBkgnd)
    TextBkgndTrans  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkBkgndTrans)
    TextDirection  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkDirection)
    TopMargin  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkTopMargin)
    VerticalAlign  = SRC(__SEC__.visSectionObject, __ROW__.visRowText, __CEL__.visTxtBlkVerticalAlign)

    #  Action tags
    SmartTags_ButtonFace  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagButtonFace)
    SmartTags_Description  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagDescription)
    SmartTags_Disabled  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagDisabled)
    SmartTags_DisplayMode  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagDisplayMode)
    SmartTags_TagName  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagName)
    SmartTags_X  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagX)
    SmartTags_XJustify  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagXJustify)
    SmartTags_Y  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagY)
    SmartTags_YJustify  = SRC(__SEC__.visSectionSmartTag, __ROW__.visRowSmartTag, __CEL__.visSmartTagYJustify)

    #  style
    EnableFillProps  = SRC(__SEC__.visSectionObject, __ROW__.visRowStyle, __CEL__.visStyleIncludesFill)
    EnableLineProps  = SRC(__SEC__.visSectionObject, __ROW__.visRowStyle, __CEL__.visStyleIncludesLine)
    EnableTextProps  = SRC(__SEC__.visSectionObject, __ROW__.visRowStyle, __CEL__.visStyleIncludesText)
    HideText  = SRC(__SEC__.visSectionObject, __ROW__.visRowStyle, __CEL__.visStyleHidden)

    # tabs
    Tabs_Alignment  = SRC(__SEC__.visSectionTab, __ROW__.visRowTab, __CEL__.visTabAlign)
    Tabs_Position  = SRC(__SEC__.visSectionTab, __ROW__.visRowTab, __CEL__.visTabPos)
    Tabs_StopCount  = SRC(__SEC__.visSectionTab, __ROW__.visRowTab, __CEL__.visTabStopCount)

    #  shape layout
    ConFixedCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOConFixedCode)
    ConLineJumpCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOJumpCode)
    ConLineJumpDirX  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOJumpDirX)
    ConLineJumpDirY  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOJumpDirY)
    ConLineJumpStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOJumpStyle)
    ConLineRouteExt  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOLineRouteExt)
    ShapeFixedCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOFixedCode)
    ShapePermeablePlace  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPermeablePlace)
    ShapePermeableX  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPermX)
    ShapePermeableY  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPermY)
    ShapePlaceFlip  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPlaceFlip)
    ShapePlaceStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPlaceStyle)
    ShapePlowCode  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOPlowCode)
    ShapeRouteStyle  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLORouteStyle)
    ShapeSplit  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOSplit)
    ShapeSplittable  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLOSplittable)
    DisplayLevel  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLODisplayLevel) #  new in Visio 2010
    Relationships  = SRC(__SEC__.visSectionObject, __ROW__.visRowShapeLayout, __CEL__.visSLORelationships) #  new in Visio 2010
