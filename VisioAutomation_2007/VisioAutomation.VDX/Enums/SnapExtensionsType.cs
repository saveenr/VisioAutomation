namespace VisioAutomation.VDX.Enums
{
    [System.Flags]
    public enum SnapExtensionsType
    {
        Nothing = 0,
        AlignmentBox = 1,
        CenterAxis = 2,
        CurveTangent = 4,
        EndPoint = 8,
        MidPoint = 16,
        Linear = 32,
        Curve = 64,
        EndPointPerpendicular = 128,
        MidPointPerpendicular = 256,
        EndPointHorizontal = 512,
        EndPointVertical = 1024,
        EllipseCenter = 2048,
        IsometricAngles = 4096,
    }
}