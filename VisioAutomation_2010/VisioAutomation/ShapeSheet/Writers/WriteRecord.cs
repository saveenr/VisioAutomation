using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal struct WriteRecord<T>
    {
        public readonly T Coord;
        public readonly string Value;
        public WriteRecord(T coord, string value)
        {
            this.Coord = coord;
            this.Value = value;
        }
    }
}