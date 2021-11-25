﻿using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class MsaglOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public VA.Geometry.Size PageBorderWidth { get; set; }
        public VA.Geometry.Size DefaultShapeSize { get; set; }
        public MsaglDirection Direction { get; set; }

        public MsaglOptions() 
        {
            this.UseDynamicConnectors = true;
            this.ScalingFactor = 14;
            this.PageBorderWidth = new VA.Geometry.Size(0.5, 0.5);
            this.DefaultShapeSize = new VA.Geometry.Size(1.0, 0.75);
            this.Direction = MsaglDirection.TopToBottom;
        }
    }
}