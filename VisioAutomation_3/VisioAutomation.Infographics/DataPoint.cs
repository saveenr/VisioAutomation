using System;

namespace VisioAutomation.Infographics
{
    public class DataPoint
    {
        // Relating to the value
        public Double Value;

        // Relating to the Label
        public string Label;

        public DataPoint(double v, string label)
        {
            this.Value = v;
            this.Label = label;
        }
    }
}