using CONTMODEL = VisioAutomation.Models.ContainerLayout;

namespace VisioAutomationSamples
{
    public static class ContainerLayoutSamples
    {
        public static void SimpleContainer()
        {
            var m = new CONTMODEL.ContainerLayout();

            var c1 = m.AddContainer("Container 1");
            var c2 = m.AddContainer("Container 2");

            c1.Add("A");

            c1.Add("B");
            c1.Add("C");

            c2.Add("1");
            c2.Add("2");
            c2.Add("3");

            m.LayoutOptions = new CONTMODEL.LayoutOptions();
            m.LayoutOptions.ContainerFormatting.FormatCells.FillForegnd = "rgb(0,176,240)";
            m.LayoutOptions.ContainerItemFormatting.FormatCells.FillForegnd = "rgb(250,250,250)";
            m.LayoutOptions.ContainerItemFormatting.FormatCells.LinePattern= "0";

            m.PerformLayout();
            m.Render(SampleEnvironment.Application.ActiveDocument);
        }
    }
}