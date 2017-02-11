using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class WriteRecords
    {
        private readonly List<WriteRecord> Records;

        private int CountSRC;
        private int CountSIDSRC;

        public WriteRecords()
        {
            this.Records = new List<WriteRecord>();
            this.CountSIDSRC = 0;
            this.CountSRC = 0;
        }

        public void Clear()
        {
            this.Records.Clear();
            this.CountSIDSRC = 0;
            this.CountSRC = 0;
        }

        public void Add(SRC src, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(src, value, unitcode);
            this.Records.Add(rec);

            this.CountSRC++;
        }

        public void Add(SIDSRC sidsrc, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(sidsrc, value, unitcode);
            this.Records.Add(rec);

            this.CountSIDSRC++;
        }

        public int Count => this.Records.Count;

        public int CountByCoordType(CoordType type)
        {
            if (this.Records.Count != (this.CountSRC + this.CountSIDSRC))
            {
                throw new System.ArgumentException();
            }

            return type == CoordType.SIDSRC ? this.CountSIDSRC : this.CountSRC;
        }

        public IEnumerable<WriteRecord> EnumerateByCoordType(CoordType type)
        {
            return this.Records.Where(i => i.Type == type);
        }
    }
}