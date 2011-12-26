using VisioAutomation.DOM;
using VisioAutomation.Format;
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
            m.LayoutOptions.Style = VA.Layout.ContainerLayout.RenderStyle.UseVisioContainers;
            m.LayoutOptions.ContainerFormatting.ShapeFormatCells.FillForegnd = "rgb(0,176,240)";
            m.LayoutOptions.ContainerItemFormatting.ShapeFormatCells.FillForegnd = "rgb(250,250,250)";
            m.LayoutOptions.ContainerItemFormatting.ShapeFormatCells.LinePattern= "0";

            m.PerformLayout();
            m.Render(SampleEnvironment.Application.ActiveDocument);

        }

    }
}