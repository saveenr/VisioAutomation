namespace VisioAutomation.Models.Layouts.Container;

public class LayoutOptions
{
    public string ManualItemStencil = "basic_u.vss";
    public string ManualItemMaster = "Rounded Rectangle";
    public string ManualContainerMaster = "Rectangle";
    public string ContainerMaster = "Container 1";

    public double ItemWidth { get; set; }
    public double ItemHeight { get; set; }
    public double ItemVerticalSpacing { get; set; }

    public double Padding { get; set; }

    public double ContainerHeaderHeight { get; set; }
    public double ContainerHorizontalDistance { get; set; }

    public Formatting ContainerFormatting { get; set; }
    public Formatting ContainerItemFormatting { get; set; }

    public LayoutOptions()
    {
        this.ContainerHeaderHeight = 0.25;
        this.Padding = 0.125;
        this.ItemVerticalSpacing = 0.125;
        this.ItemHeight = 0.25;
        this.ContainerHorizontalDistance = 1.0;
        this.ItemWidth = 2.0;
        this.ContainerFormatting = new Formatting();
        this.ContainerItemFormatting = new Formatting();
        this.ContainerFormatting.TextBlockCells.VerticalAlign = "0";
    }
}