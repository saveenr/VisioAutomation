using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Charting
{
    public class DataPoint
    {
        public double Value;
        public string Label;
        public string LabelFormat;
        public IVisio.Shape VisioShape;

        public DataPoint(double value)
        {
            this.Value = value;
        }

        public static List<DataPoint> DoublesToDataPoints(double[] values, string[] labels)
        {
            var datapoints = new List<DataPoint>();

            for (int i = 0; i < values.Length; i++)
            {
                var dp = new DataPoint(values[i]);
                if (labels != null && i < labels.Length)
                {
                    dp.Label = labels[i];
                }

                datapoints.Add(dp);
            }
            return datapoints;
        }

    }
}