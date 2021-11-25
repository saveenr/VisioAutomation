using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Core
{
    public struct ShapeIDPair
    {
        public readonly IVisio.Shape Shape;
        public readonly int ShapeID;

        public ShapeIDPair(IVisio.Shape shape)
        {
            this.Shape = shape;
            this.ShapeID = shape.ID16;
        }
    }

}

