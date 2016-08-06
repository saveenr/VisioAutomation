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

        public static List<DataPoint> DoublesToDataPoints(double[] Values, string[] Labels)
        {
            var datapoints = new List<DataPoint>();

            for (int i = 0; i < Values.Length; i++)
            {
                var dp = new DataPoint(Values[i]);
                if (Labels != null && i < Labels.Length)
                {
                    dp.Label = Labels[i];
                }

                datapoints.Add(dp);
            }
            return datapoints;
        }

    }
}