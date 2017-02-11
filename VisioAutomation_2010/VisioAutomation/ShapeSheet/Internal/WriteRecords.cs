using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class WriteRecords
    {
        private readonly List<WriteRecord> _records;

        private int CountSRC;
        private int CountSIDSRC;

        public WriteRecords()
        {
            this._records = new List<WriteRecord>();
            ResetCounts();
        }

        public WriteRecords(int capacity)
        {
            this._records = new List<WriteRecord>(capacity);
            ResetCounts();
        }

        public void Clear()
        {
            this._records.Clear();
            ResetCounts();
        }

        private void ResetCounts()
        {
            this.CountSIDSRC = 0;
            this.CountSRC = 0;
        }

        public void Add(SRC src, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(src, value, unitcode);
            this._records.Add(rec);
            this.CountSRC++;
        }

        public void Add(SIDSRC sidsrc, string value, IVisio.VisUnitCodes? unitcode)
        {
            var rec = new WriteRecord(sidsrc, value, unitcode);
            this._records.Add(rec);
            this.CountSIDSRC++;
        }

        public int Count => this._records.Count;

        public int CountByCoordType(CoordType type)
        {
            if (this._records.Count != (this.CountSRC + this.CountSIDSRC))
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Mismatch in counting number of records");
            }

            return type == CoordType.SIDSRC ? this.CountSIDSRC : this.CountSRC;
        }

        public IEnumerable<WriteRecord> EnumerateByCoordType(CoordType type)
        {
            return this._records.Where(i => i.CoordType == type);
        }
    }
}