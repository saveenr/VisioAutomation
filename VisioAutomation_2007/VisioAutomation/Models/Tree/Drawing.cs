using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Tree
{
    public class Drawing
    {
        public Node Root { get; set; }
        public VA.Models.Tree.LayoutOptions LayoutOptions;
        
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
            get { return VA.Internal.TreeOps.PreOrder(this.Root, n => n.Children); }
        }
    }
}