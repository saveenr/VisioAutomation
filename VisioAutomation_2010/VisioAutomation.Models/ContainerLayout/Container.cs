using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.ContainerLayout
{
    public class Container
    {
        public VisioAutomation.Models.Text.TextElement Text { get; set; }
        public List<ContainerItem> ContainerItems { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public Drawing.Rectangle Rectangle;
        public short ShapeID;

        public Container(VisioAutomation.Models.Text.TextElement text)
        {
            this.Text = text;
            this.ContainerItems = new List<ContainerItem>();
        }

        public Container(string text) :
            this( new VisioAutomation.Models.Text.TextElement(text))
        {
        }

        public ContainerItem Add(string text)
        {
            var ct = new ContainerItem(text);
            this.ContainerItems.Add(ct);
            return ct;
        }
    }
}
