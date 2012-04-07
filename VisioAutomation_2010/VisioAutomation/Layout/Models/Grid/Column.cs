using System;

namespace VisioAutomation.Layout.Models.Grid
{
    public class Column
    {
        private double _width;
        public object Data { get; set; }

        public double Width
        {
            get { return _width; }
            set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException();
                }
                _width = value;
            }
        }
    }
}