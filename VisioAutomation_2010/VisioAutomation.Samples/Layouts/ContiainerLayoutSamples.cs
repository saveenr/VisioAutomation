using VisioAutomation.Models.Layouts.Container;

namespace VisioAutomationSamples;

public static class ContainerLayoutSamples
{
    public static void SimpleContainer()
    {
        var m = new ContainerLayout();

        var c1 = m.AddContainer("Container 1");
        var c2 = m.AddContainer("Container 2");

        c1.Add("A");

        c1.Add("B");
        c1.Add("C");

        c2.Add("1");
        c2.Add("2");
        c2.Add("3");

        m.LayoutOptions = new LayoutOptions();
        m.LayoutOptions.ContainerFormatting.ShapeFormatCells.FillForeground = "rgb(0,176,240)";
        m.LayoutOptions.ContainerItemFormatting.ShapeFormatCells.FillForeground = "rgb(250,250,250)";
        m.LayoutOptions.ContainerItemFormatting.ShapeFormatCells.LinePattern= "0";

        m.PerformLayout();
        m.Render(SampleEnvironment.Application.ActiveDocument);
    }
}