using System;

namespace VisioAutomation.Models.Layouts.Grid;

public class Row
{
    private double _height;
    public object Data { get; set; }

    public double Height
    {
        get { return this._height; }
        set
        {
            if (value <= 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            this._height = value;
        }
    }
}