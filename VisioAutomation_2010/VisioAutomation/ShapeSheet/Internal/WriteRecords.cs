using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class WriteRecords
    {
        private readonly List<WriteRecord> Records;

        public WriteRecords()
        {
            this.Records = new List<WriteRecord>();
        } 

        public void Clear()
        {
            this.Records.Clear();
        }

        public void Add(SRC src, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(src, value, unitcode);
            this.Records.Add(rec);
        }

        public void Add(SIDSRC sidsrc, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(sidsrc, value, unitcode);
            this.Records.Add(rec);
        }

        public int Count => this.Records.Count;

        public IEnumerable<WriteRecord> EnumerateByCoordType(CoordType type)
        {
            return this.Records.Where(i => i.Type == type);
        }
    }
}