
# Use SnagIt or Hypersnap to capture the text of a Visio ShapeSheet window
# this script will approximately draw the shapesheet cells into a new visio digram

#example input it below

inputtext="""

  Page Properties
	PageWidth	8.5 in	PageScale	1 in	ShdwOffsetX	0.125 in
	PageHeight	11 in	DrawingScale	1 in	ShdwOffsetY	-0.125 in
	DrawingSizeType	0	DrawingScaleType	0	InhibitSnap	FALSE
	ShdwType  0	ShdwObliqueAngle	0 deg	ShdwScaleFactor  100%
	UIVisibility 0
  Page Layout
	PlaceStyle	0	BlockSizeX	0.25 in	LineToNodeX	0.125 in
	PlaceDepth  0	BlockSizeY	0.25 in	LineToNodeY	0.125 in
	PlowCode	0	AvenueSizeX	0.375 in	LineToLineX  0.125 in
	ResizePage  FALSE	AvenueSizeY	0.375 in	LineToLineY  0.125 in
	DynamicsOff	FALSE	RouteStyle  0	LineJumpFactorX	0.6667
	EnableGrid FALSE	PageLineJumpDirX	0	LineJumpFactorY	0.6667
	CtrlAsInput FALSE	PageLineJumpDirY	0	LineJumpCode	1
	LineAdjustFrom  0	LineAdjustTo  0	LineJumpStyle  0
	PlaceFlip  0	LineRouteExt  0	PageShapeSplit  1
  Ruler & Grid
	XRulerOrigin	0 in	XGridOrigin	0 in	XGridSpacing	0 in
	YRulerOrigin	0 in	YGridOrigin	0 in	YGridSpacing	0 in
	XRulerDensity  32	XGridDensity 8
	YRulerDensity  32	YGridDensity 8
  Annotation	X	Y	MarkerIndex	Date	Comment	LangID
  1	4.83 in	7.795 in	1	DATETIME(40659.1753)	""	1045
  2	17 in	9 in	2	DATETIME(40659.1758)	"xcc"	1045
  Print Properties
	PageLeftMargin	0.25 in	PageRightMargin	0.25 in
	PageTopMargin	0.25 in	PageBottomMargin	0.25 in
	ScaleX 100%	ScaleY 100%
	PagesX  1	PagesY  1
	CenterX  FALSE	CenterY  FALSE
	OnPage  FALSE	PrintGrid  FALSE
	PrintPageOrientation 1	PaperKind  1
	PaperSource	7





 Shape Transform
	Width  4 in	PinX  15 in	FlipX FALSE
	Height  4 in	PinY  7 in	FlipY FALSE
	Angle 0 deg	LocPinX Width*0.5	ResizeMode 0
	LocPinY Height*0.5
 Shape Data	Label	Prompt	Type	Format	Value	SortKey	Invisible	Ask	LangID	Calendar
	Prop.PROP1 "Property1"	No Formula	0	No Formula	"TESTVALUE"	No Formula	No Formula	No Formula	1045	No Formula
 Hyperlinks	Description	Address	SubAddress	ExtraInfo	Frame	SortKey	NewWindow	Default	Invisible
	Hyperlink.Row_1 "microsoft"	"http://microsoft.com"	""	""	""	""	FALSE	FALSE	FALSE
 Geometry 1
	Geometry1.NoFill FALSE	Geometry1.NoLine FALSE	Geometry1.NoShow FALSE	Geometry1.NoSnap FALSE
	Name	X	Y	A	B	C	D	E
 1	MoveTo Width*0	Height*0
 2	LineTo Width*1	Height*0
 3	LineTo Width*1	Height*1
 4	LineTo Width*0	Height*1
 5	LineTo Geometry1.X1	Geometry1.Y1
 Protection
	LockWidth 0	LockEnd 0	LockCrop 0
	LockHeight 0	LockDelete 0	LockGroup 0
	LockAspect 0	LockSelect 0	LockCalcWH 0
	LockMoveX 0	LockFormat 0	LockFromGroupFormat 0
	LockMoveY 0	LockCustProp 0	LockThemeColors 0
	LockRotate 0	LockTextEdit 0	LockThemeEffects 0
	LockBegin 0	LockVtxEdit 0
 Miscellaneous
	NoObjHandles FALSE	HideText FALSE	ObjType 0
	NoCtlHandles FALSE	UpdateAlignBox FALSE	IsDropSource FALSE
	NoAlignBox FALSE	DynFeedback 0	Comment "screentip"
	NonPrinting FALSE	NoLiveDynamics FALSE	DropOnPageScale 100%
	LangID 1033	Calendar 0	LocalizeMerge FALSE
 Line Format
	LinePattern 1	BeginArrow 0	BeginArrowSize 2
	LineWeight 0.72 pt	EndArrow 0	EndArrowSize 2
	LineColor 0	LineColorTrans 0%	Rounding  0 in
	LineCap 0
 Fill Format
	FillForegnd 1	ShdwForegnd 0	ShapeShdwOffsetX  0 in
	FillForegndTrans 0%	ShdwForegndTrans 0%	ShapeShdwOffsetY  0 in
	FillBkgnd 0




 Text Fields	Format	Value	Calendar	ObjectKind
	0	FIELDPICTURE(0)	Prop.PROP1	0	0
 Character	Font	Size	Scale	Spacing	Color	Transparency	Style	Case	Pos.	Strikethru	DoubleULine	Overline	DoubleStrikethrough	AsianFont	ComplexScriptFont	LocalizeFont	ComplexScriptSize	LangID
	3	4	12 pt	100%	0 pt	0	0%	0	0	0	FALSE	FALSE	FALSE	FALSE	0	0	0	-100%	1045
	3	4	12 pt	100%	0 pt	0	0%	0	0	0	FALSE	FALSE	FALSE	FALSE	0	0	0	-100%	1045
 Paragraph	IndFirst	IndLeft	IndRight	SpLine	SpBefore	SpAfter	HAlign	Bullet	BulletString	BulletFont	LocBulletFont TextPosAfterBullet	BulletSize	Flags
	0	0 in	0 in	0 in	-120%	0 pt	0 pt	1	0	""	0	0	0 in	-100%	0
 Tabs	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignment	Position	Alignme
	0
 Text Block Format
	LeftMargin  4 pt	TopMargin	4 pt	TextDirection	0
	RightMargin	4 pt	BottomMargin	4 pt	VerticalAlign	1
	TextBkgnd	0	TextBkgndTrans  0%	DefaultTabStop	0.5 in
 Events
	TheData	No Formula	EventDblClick	No Formula	EventDrop	No Formula
	TheText	No Formula	EventXFMod No Formula	EventMultiDrop  No Formula
 Image Properties
	Contrast 50%	Gamma 1	Sharpen  0%
	Brightness 50%	Blur 0%	Denoise 0%
	Transparency  0%
 Glue Info
	BegTrigger  No Formula	GlueType  0	WalkPreference	0
	EndTrigger  No Formula
 Shape Layout
  ShapePermeableX	FALSE	ShapePermeableY	FALSE	ShapePermeablePlace	FALSE
	ShapeFixedCode	0	ShapePlowCode	0	ShapeRouteStyle	0
	ConLineJumpDirX	0	ConLineJumpDirY	0	ConFixedCode	0
 ConLineJumpCode	0	ConLineJumpStyle	0	ShapeSplit 0
	ShapePlaceFlip 0	ConLineRouteExt	0	ShapeSplittable	0
	ShapePlaceStyle  0



"""

import sys
f = file("d:\\shapesheet.txt")
lines = f.readlines()
f.close()

lines = [ s.strip() for s in lines]

import sys
import clr
import System
clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio
IVisio = Microsoft.Office.Interop.Visio
visapp = IVisio.ApplicationClass()
doc = visapp.Documents.Add("")
page = visapp.ActivePage

cy =0;

cellwidth = 1.5;
cellheight = 0.5
cellvsep = 0.0
cellhsep = 0.0

for line in lines :
    cx=0;
    tokens0 = line.split()
    tokens = []
    for t in tokens0 :
        if (t in [ "in" , "pt" , "deg"]) :
            n = len(tokens)-1
            tokens[n] = tokens[n] + " " + t
        if (t == "Formula") :
            n = len(tokens)-1
            if (tokens[n]=="No") :
                tokens[n] = tokens[n] + " " + t
            else :
                tokens.append(t)
        else :
            tokens.append(t)
        
    for t in tokens :
        shape = page.DrawRectangle(cx,cy,cx+cellwidth,cy-cellheight)
        shape.Text = t
        cx = cx + cellwidth
    cy = cy - cellheight - cellvsep
        
page.ResizeToFitContents()

    