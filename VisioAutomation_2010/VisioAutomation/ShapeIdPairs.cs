using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public class ShapeIdPairs : List<ShapeIdPair>
    {
        public ShapeIdPairs()
        {
        }

        public ShapeIdPairs(int capacity) : base (capacity)
        {
        }

        public static ShapeIdPairs FromShapes(IList<IVisio.Shape> shapes)
        {
            var shapeidpairs = new ShapeIdPairs(shapes.Count);
            shapeidpairs.AddRange(shapes.Select(s => new ShapeIdPair(s)));
            return shapeidpairs;
        }

        public static ShapeIdPairs FromShapes(params IVisio.Shape[] shapes)
        {
            var shapeidpairs = new ShapeIdPairs(shapes.Length);
            shapeidpairs.AddRange(shapes.Select(s => new ShapeIdPair(s)));
            return shapeidpairs;
        }
    }
}
