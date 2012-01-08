using VisioAutomation.DOM;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BoxL = VisioAutomation.Layout.BoxLayout2;

namespace VisioAutomationSamples
{
    public static class BoxLayout2Samples
    {
        public static void BoxLayout_1()
        {
            // Create a layout
            var layout = new VA.Layout.BoxLayout2.BoxLayout();

            var root = layout.Root;
            root.MinHeight = 10;
            root.Direction = BoxL.ContainerDirection.Vertical;
            root.AddBox(1,2);
            root.AddBox(1,1);

            layout.PerformLayout();

            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);
            var page1 = doc.Pages[1];

            // and tinker with it
            // render
            var nodes = layout.Nodes.ToList();
            foreach (var node in nodes)
            {
                page1.DrawRectangle(node.Rectangle);
            }

            var margin = new VA.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);
        }

    }
}