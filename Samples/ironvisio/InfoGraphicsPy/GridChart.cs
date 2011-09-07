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
    public class GridChart : Chart
    {
        private double cellwidth = 0.5;
        public double HorizontalSeparation = 0.10;
        public double VerticalSeparation = 0.10;
        public double CellHeight = 0.5;
        public double CategoryLabelHeight = 0.5;
        public double CellWidth
        {
            get { return cellwidth; }
            set { cellwidth = value; }
        }


        public string LineLightBorder = "rgb(220,220,220)";
        public string ValueFillColor = "rgb(240,240,240)";
        public string NonValueColor = "rgb(255,255,255)";
        public string CategoryFillPattern = "0";
        public string CategoryLineWeight = "0.0";
        public string CategoryLinePattern = "0";

        
    }
}
