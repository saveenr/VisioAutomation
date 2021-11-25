

namespace VisioAutomation.Models.Layouts.Container;

public class ContainerItem
{
    public string Text { get; set; }
    public VisioAutomation.Geometry.Rectangle Rectangle { get; set; }
    public IVisio.Shape VisioShape { get; set; }
    public short ShapeID { get; set; }
        
    public ContainerItem(string text)
    {
        this.Text = text;
    }
}