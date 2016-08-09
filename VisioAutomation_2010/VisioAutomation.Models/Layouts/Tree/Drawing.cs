using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.Tree
{
    public class Drawing
    {
        public Node Root { get; set; }
        public LayoutOptions LayoutOptions;
        
        public Drawing()
        {
            this.LayoutOptions = new LayoutOptions();            
        }

        public void Render(IVisio.Page page)
        {
            var renderer = new TreeLayout();
            if (this.LayoutOptions != null)
            {
                renderer.LayoutOptions = this.LayoutOptions;                
            }
            renderer.RenderToVisio(this, page);
        }
        
        public IEnumerable<Node> Nodes
        {
            get { return GenTreeOps.Algorithms.PreOrder(this.Root, n => n.Children); }
        }
    }
}