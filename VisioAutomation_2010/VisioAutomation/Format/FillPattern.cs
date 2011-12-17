using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public enum FillPattern
    {
        None = 0,
        Solid = IVisio.VisCellVals.visSolid,
        WideUpDiagonal = IVisio.VisCellVals.visWideUpDiagonal,
        WideCross = IVisio.VisCellVals.visWideCross,
        WideDiagonalCross = IVisio.VisCellVals.visWideDiagonalCross,
        WideDownDiagonal = IVisio.VisCellVals.visWideDownDiagonal,
        WideHorz = IVisio.VisCellVals.visWideHorz,
        WideVert = IVisio.VisCellVals.visWideVert,
        BackDotsMini = IVisio.VisCellVals.visBackDotsMini,
        HalfAndHalf = IVisio.VisCellVals.visHalfAndHalf,
        ForeDotsMini = IVisio.VisCellVals.visForeDotsMini,
        ForeDotsNarrow = IVisio.VisCellVals.visForeDotsNarrow,
        ForeDotsWide = IVisio.VisCellVals.visForeDotsWide,
        ThickHorz = IVisio.VisCellVals.visThickHorz,
        ThickVertical = IVisio.VisCellVals.visThickVertical,
        ThickDownDiagonal = IVisio.VisCellVals.visThickDownDiagonal,
        ThickUpDiagonal = IVisio.VisCellVals.visThickUpDiagonal,
        ThickDiagonalCross = IVisio.VisCellVals.visThickDiagonalCross,
        BackDotsWide = IVisio.VisCellVals.visBackDotsWide,
        ThinHorz = IVisio.VisCellVals.visThinHorz,
        ThinVert = IVisio.VisCellVals.visThinVert,
        ThinDownDiagonal = IVisio.VisCellVals.visThinDownDiagonal,
        ThinUpDiagonal = IVisio.VisCellVals.visThinUpDiagonal,
        ThinCross = IVisio.VisCellVals.visThinCross,
        ThinDiagonalCross = IVisio.VisCellVals.visThinDiagonalCross,
        LinearLeftToRight = 25,
        LinearVertical = 26,
        LinearRightToLeft = 27,
        LinearTopToBottom = 28,
        LinearHorizontal = 29,
        LinearBottomToTop = 30,
        RectangularUpperLeft = 31,
        RectangularUpperRight = 32,
        RectangularLowerLeft = 33,
        RectangularLowerRight = 34,
        RectangularCenter = 35,
        RadialUpperLeft = 36,
        RadialUpperRight = 37,
        RadialLowerLeft = 38,
        RadialLowerRight = 39,
        RadialCenter = 40
    }
}