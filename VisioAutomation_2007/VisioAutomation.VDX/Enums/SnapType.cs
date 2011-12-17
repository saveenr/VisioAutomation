namespace VisioAutomation.VDX
{
    [System.Flags]
    public enum SnapType
    {
        Nothing = 0,
        RuleSubdivisions = 1,
        Grid = 2,
        Guides = 4,
        SelectionHandles = 8,
        Vertices = 16,
        ConnectionPoints = 32,
        VisibleShapEdges = 256,
        AlignmentBox = 512,
        ShapeExtensions = 1024,
        Disabled = 32768,
        Intersetions = 65536
    }
}