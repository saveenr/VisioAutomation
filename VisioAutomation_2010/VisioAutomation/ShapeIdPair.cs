using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
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

        public ShapeIDPair(IVisio.Shape shape, int id)
        {
            this.Shape = shape;
            this.ShapeID = id;
        }
    }

}

