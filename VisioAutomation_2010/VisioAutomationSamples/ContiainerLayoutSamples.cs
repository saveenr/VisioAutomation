using VisioAutomation.DOM;
using VisioAutomation.Layout.ContainerLayout;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BH = VisioAutomation.Layout.BoxLayout;

namespace VisioAutomationSamples
{
    public static class ContainerLayoutSamples
    {
        public static void SimpleContainer()
        {

            var cont_fmt = new VA.Format.ShapeFormatCells();
            cont_fmt.FillForegnd = "rgb(150,180,240)";
            cont_fmt.FillBkgnd= "rgb(150,180,240)";
            cont_fmt.FillPattern = "40";
            cont_fmt.LinePattern = "0";
            var cont_tb = new VA.Text.TextBlockFormatCells();
            cont_tb.VerticalAlign = "0";

            var cont_char = new VA.Text.CharacterFormatCells();
            cont_char.Font= "10";

            var m = new VA.Layout.ContainerLayout.ContainerLayout();

            var c1 = m.AddContainer("Container 1");
            var c2 = m.AddContainer("Container 2");

            c1.Add("A");

            c1.Add("B");
            c1.Add("C");

            c2.Add("1");
            c2.Add("2");
            c2.Add("3");

            m.LayoutOptions = new LayoutOptions();
            m.LayoutOptions.Style = VA.Layout.ContainerLayout.RenderStyle.UseShapes;
            m.PerformLayout();
            m.Render(SampleEnvironment.Application.ActiveDocument);

        }

    }
}