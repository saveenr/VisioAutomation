using System.Collections.Generic;

namespace VisioAutomation.Models.Charting
{
    public class DataPointList : IEnumerable<DataPoint>
    {
        private readonly List<DataPoint> _items;

        public DataPointList()
        {
            this._items = new List<DataPoint>();
        }

        public DataPointList(IList<double> doubles, IList<string> labels)
        {
            this._items = new List<DataPoint>(doubles.Count);
            for (int i = 0; i < doubles.Count; i++)
            {
                var dp = new DataPoint(doubles[i]);
                if (labels != null && i < labels.Count)
                {
                    dp.Label = labels[i];
                }

                this._items.Add(dp);
            }
        }

        public int Count
        {
            get
            {
                return this._items.Count;
            }
        }

        public IEnumerator<DataPoint> GetEnumerator()
        {
            foreach (var i in this._items)
            {
                yield return i;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public DataPoint this[int index]
        {
            get { return this._items[index]; }
        }

        public void Add(double d)
        {
            var dp = new DataPoint(d);
            this._items.Add(dp);
        }

        public void Add(DataPoint dp)
        {
            this._items.Add(dp);
        }
    }
}