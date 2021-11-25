using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.Container
{
    public class Container
    {
        public VisioAutomation.Models.Text.Element Text { get; set; }
        public List<ContainerItem> ContainerItems { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public VisioAutomation.Geometry.Rectangle Rectangle;
        public short ShapeID;

        public Container(VisioAutomation.Models.Text.Element text)
        {
            this.Text = text;
            this.ContainerItems = new List<ContainerItem>();
        }

        public Container(string text) :
            this( new VisioAutomation.Models.Text.Element(text))
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
