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
    public class DataPoint
    {
        public double Value;
        public string Text;
        public string Tooltip;

        public DataPoint(double v)
        {
            this.Value = v;
            this.Text = null;
            this.Tooltip = null;
        }

        public DataPoint(double v, string t)
        {
            this.Value = v;
            this.Text = t;
            this.Tooltip = null;
        }

    }
}
