using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class MsaglOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        private VA.Core.Size? _page_border_width;

        /// <summary>
        /// Margin around the drawing when the page is resized to fit it, in document units (inches).
        /// Until it is set, 0.5 x 0.5 applies when the renderer is used directly, and
        /// DrawDirectedGraphDocument uses the document's BorderSize instead. Setting it, even to 0.5 x 0.5,
        /// makes it win over the document's BorderSize for the page drawn from this layout.
        /// </summary>
        public VA.Core.Size PageBorderWidth
        {
            get { return this._page_border_width ?? new VA.Core.Size(0.5, 0.5); }
            set { this._page_border_width = value; }
        }

        /// <summary>
        /// True once PageBorderWidth has been set on this instance.
        /// </summary>
        public bool IsPageBorderWidthSet
        {
            get { return this._page_border_width.HasValue; }
        }
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
            this.DefaultShapeSize = new VA.Core.Size(1.0, 0.75);
            this.Direction = MsaglDirection.TopToBottom;
            this.EdgeLabelBoxSize = new VA.Core.Size(1.0, 0.5);
        }
    }
}