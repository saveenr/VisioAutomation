using System;

namespace VisioAutomation.Models.Grid
{
    public class Row
    {
        private double _height;
        public object Data { get; set; }

        public double Height
        {
            get { return _height; }
            set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException();
                }
                _height = value;
            }
        }
    }
}