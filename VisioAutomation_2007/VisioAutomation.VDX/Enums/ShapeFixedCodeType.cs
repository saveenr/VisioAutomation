namespace VisioAutomation.VDX.Enums
{
    [System.Flags]
    public enum ShapeFixedCodeType
    {
        None = 0,
        DontMoveOnLayOutShapesCommand = 1,
        DontMoveOnDisallowPlaceable = 2,
        DontMoveOnAllowPlaceable = 4,
        IgnoreConnectionPoints = 32,
        OnlyAllowRoutingToSides = 64,
        GlueToAlignmentBox = 128
    }
}