using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
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
