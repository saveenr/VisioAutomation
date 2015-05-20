using System;
using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    public class DataPoints : IEnumerable<DataPoint>
    {
        private List<DataPoint> points;

        public DataPoints()
        {
            this.points = new List<DataPoint>();
        }

        public DataPoints(IList<double> values)
        {
            this.points = new List<DataPoint>(values.Count);
            foreach (double v in values)
            {
                this.Add(v);
            }
        }

        public IEnumerator<DataPoint> GetEnumerator()
        {
            foreach (var i in this.points)
                yield return i;
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }

        public DataPoint Add(double value)
        {
            var dp = new DataPoint(value,value.ToString());
            dp.Value = value;
            this.points.Add(dp);
            return dp;
        }

        public DataPoint this[int index]
        {
            get { return this.points[index]; }
        }

        public List<double> GetNormalizedValues(double s)
        {
            double max = this.Select(dp => dp.Value).Max();
            var items = new List<double>(this.Count);
            foreach (var dp in this)
            {
                items.Add((dp.Value/max)*s);
            }
            return items;
        }

        public List<double> GetNormalizedValues()
        {
            return this.GetNormalizedValues(1.0);
        }

        public int Count
        {
            get { return this.points.Count; }
        }
    }
}
