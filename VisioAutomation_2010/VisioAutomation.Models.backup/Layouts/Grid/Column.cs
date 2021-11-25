using System;

namespace VisioAutomation.Models.Layouts.Grid
{
    public class Column
    {
        private double _width;
        public object Data { get; set; }

        public double Width
        {
            get { return this._width; }
            set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException();
                }
                this._width = value;
            }
        }
    }
}