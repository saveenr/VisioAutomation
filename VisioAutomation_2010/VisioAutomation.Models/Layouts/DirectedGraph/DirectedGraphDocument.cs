using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphDocument
    {
        public readonly List<DirectedGraphLayout> Layouts;
        public string Template;
        /// <summary>Margin around the drawing on each page. A layout whose LayoutOptions.PageBorderWidth has been set uses that instead.</summary>
        public VisioAutomation.Core.Size BorderSize;

        public DirectedGraphDocument()
        {
            this.Layouts = new List<DirectedGraphLayout>();
            this.Template = null;
            this.BorderSize = new VisioAutomation.Core.Size(1.0, 1.0);
        }
    }
}