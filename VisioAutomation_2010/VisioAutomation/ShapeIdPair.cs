using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public struct ShapeIdPair
    {
        public readonly IVisio.Shape Shape;
        public readonly int ShapeID;

        public ShapeIdPair(IVisio.Shape shape)
        {
            this.Shape = shape;
            this.ShapeID = shape.ID16;
        }

        public ShapeIdPair(IVisio.Shape shape, int id)
        {
            this.Shape = shape;
            this.ShapeID = id;
        }
    }

}
