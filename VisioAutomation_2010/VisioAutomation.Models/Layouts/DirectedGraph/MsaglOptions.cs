using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class MsaglOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public VA.Core.Size PageBorderWidth { get; set; }
        public VA.Core.Size DefaultShapeSize { get; set; }
        public MsaglDirection Direction { get; set; }

        /// <summary>
        /// Size reserved for each edge's label when laying out the graph, in document units (inches).
        /// Reserved even for edges without a label, so smaller values give tighter layouts.
        /// </summary>
        public VA.Core.Size EdgeLabelBoxSize { get; set; }

        /// <summary>
        /// Minimum distance between layers (rows for TopToBottom, columns for LeftToRight),
        /// in document units (inches). When null, the MSAGL default is used.
        /// </summary>
        public double? LayerSeparation { get; set; }

        public MsaglOptions() 
        {
            this.UseDynamicConnectors = true;
            this.ScalingFactor = 14;
            this.PageBorderWidth = new VA.Core.Size(0.5, 0.5);
            this.DefaultShapeSize = new VA.Core.Size(1.0, 0.75);
            this.Direction = MsaglDirection.TopToBottom;
            this.EdgeLabelBoxSize = new VA.Core.Size(1.0, 0.5);
        }
    }
}