using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.ContainerLayout
{
    public class ContainerItem
    {
        public string Text { get; set; }
        public Drawing.Rectangle Rectangle { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public short ShapeID { get; set; }
        
        public ContainerItem(string text)
        {
            this.Text = text;
        }
    }
}
