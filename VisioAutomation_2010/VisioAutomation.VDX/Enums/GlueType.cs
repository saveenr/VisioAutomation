namespace VisioAutomation.VDX.Enums
{
    [System.Flags]
    public enum GlueType
    {
        Enabled = 0,
        Guides = 1,
        Handles = 2,
        Vertices = 4,
        ConnectionPoints = 8,
        Geometry = 32
    }
}