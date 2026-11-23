using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphDocument
    {
        public readonly List<DirectedGraphLayout> Layouts;
        public string Template;
        /// <summary>Margin around the drawing on each page. A layout whose LayoutOptions.PageBorderWidth has been set uses that instead.</summary>
        public VisioAutomation.Core.Size BorderSize;
        /// <summary>
        /// Styling for every layout in the document (currently the connector master and its stencil), or null for none.
        /// When set, it is used instead of the styling passed to DrawDirectedGraphDocument, so a document that carries its
        /// own styling (for example from XML) is drawn the same way however it is drawn. An edge that sets its own
        /// MasterName and StencilName still wins over both.
        /// </summary>
        public DirectedGraphStyling Styling;

        public DirectedGraphDocument()
        {
            this.Layouts = new List<DirectedGraphLayout>();
            this.Template = null;
            this.BorderSize = new VisioAutomation.Core.Size(1.0, 1.0);
        }
    }
}