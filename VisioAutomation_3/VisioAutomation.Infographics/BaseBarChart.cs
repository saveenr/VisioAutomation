using System.Collections.Generic;
using VisioAutomation.Drawing;

namespace VisioAutomation.Infographics
{
    public abstract class BaseBarChart : Block
    {
        public IList<DataPoint> DataPoints;
        public ColorRGB ValueColor = new ColorRGB(0xa0a0a0);
        public ColorRGB NonValueColor = new ColorRGB(0xffffff);
        protected double TileHeight = 3.0;
        protected Size margin = new Size(0.25, 0.25);
        protected double _labelHeight = 0.5;
        protected double _barDistance = 0.0125;
        protected double bar_thickness = 0.5;
        protected double maxval = 180.0;

        public BaseBarChart()
        {
            this.DataPoints = new List<DataPoint>();           
        }
    }
}