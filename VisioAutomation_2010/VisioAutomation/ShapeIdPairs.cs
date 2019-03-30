using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public class ShapeIDPairs : List<ShapeIDPair>
    {
        public ShapeIDPairs()
        {
        }

        public ShapeIDPairs(int capacity) : base (capacity)
        {
        }

        public static ShapeIDPairs FromShapes(IList<IVisio.Shape> shapes)
        {
            var shapeidpairs = new ShapeIDPairs(shapes.Count);
            shapeidpairs.AddRange(shapes.Select(s => new ShapeIDPair(s)));
            return shapeidpairs;
        }

        public static ShapeIDPairs FromShapes(params IVisio.Shape[] shapes)
        {
            var shapeidpairs = new ShapeIDPairs(shapes.Length);
            shapeidpairs.AddRange(shapes.Select(s => new ShapeIDPair(s)));
            return shapeidpairs;
        }
    }
}
