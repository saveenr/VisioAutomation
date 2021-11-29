using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Core
{
    public class ShapeIDPairs : List<ShapeIDPair>
    {
        private ShapeIDPairs(int capacity) : base (capacity)
        {
        }

        public static ShapeIDPairs FromShapes(IList<IVisio.Shape> shapes)
        {
            return _from_shapes(shapes);
        }

        public static ShapeIDPairs FromShapes(params IVisio.Shape[] shapes)
        {
            return _from_shapes(shapes);
        }

        private static ShapeIDPairs _from_shapes(IList<IVisio.Shape> shapes)
        {
            var shapeidpairs = new ShapeIDPairs(shapes.Count);
            shapeidpairs.AddRange(shapes.Select(s => new ShapeIDPair(s)));
            return shapeidpairs;
        }

    }
}
