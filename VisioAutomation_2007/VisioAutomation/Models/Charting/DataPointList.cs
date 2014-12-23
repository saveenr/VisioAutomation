using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    public class DataPointList : IEnumerable<DataPoint>
    {
        private List<DataPoint> items;

        public DataPointList()
        {
            this.items = new List<DataPoint>();
        }

        public DataPointList(IList<double> doubles, IList<string> labels)
        {
            this.items = new List<DataPoint>(doubles.Count);
            for (int i = 0; i < doubles.Count; i++)
            {
                var dp = new DataPoint(doubles[i]);
                if (labels != null && i < labels.Count)
                {
                    dp.Label = labels[i];
                }

                this.items.Add(dp);
            }
        }

        public int Count
        {
            get
            {
                return this.items.Count;
            }
        }

        public IEnumerator<DataPoint> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public DataPoint this[int index]
        {
            get { return this.items[index]; }
        }

        public void Add(double d)
        {
            var dp = new DataPoint(d);
            this.items.Add(dp);
        }

        public void Add(DataPoint dp)
        {
            this.items.Add(dp);
        }
    }
}