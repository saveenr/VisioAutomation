using VA=VisioAutomation;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.Tree
{
    public class Drawing
    {
        public Node Root { get; set; }
        public VA.Layout.Tree.LayoutOptions LayoutOptions;
        
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
    }
}