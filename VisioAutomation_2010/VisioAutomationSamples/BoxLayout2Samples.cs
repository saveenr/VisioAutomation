using VisioAutomation.DOM;
using VisioAutomation.Layout.BoxLayout2;
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
        public static void BoxLayout_SimpleCases()
        {
            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);

            // Create a simple Column
            var layout1 = new BoxL.BoxLayout();
            layout1.Root = new BoxL.Container( BoxL.ContainerDirection.Vertical );
            layout1.Root.Padding = 0.25;
            layout1.Root.ChildSeparation = 0.25;
            layout1.Root.AddBox(1,2);
            layout1.Root.AddBox(1,1);

            // You can set the min height and width of a container
            var layout2 = new VA.Layout.BoxLayout2.BoxLayout();
            layout2.Root = new BoxL.Container(BoxL.ContainerDirection.Vertical);
            layout2.Root.MinHeight = 5;
            layout2.Root.MinWidth = 3;
            layout2.Root.Padding = 0.25;
            layout2.Root.ChildSeparation = 0.25;
            layout2.Root.AddBox(1, 2);
            layout2.Root.AddBox(1, 1);

            // For vertical containers, you can layout shapes bottom-to-top or top-to-bottom
            var layout3 = new BoxL.BoxLayout();
            layout3.Root = new BoxL.Container(BoxL.ContainerDirection.Vertical);
            layout3.Root.MinHeight = 5;
            layout3.Root.MinWidth = 3;
            layout3.Root.ChildVerticalDirection = DirectionVertical.TopToBottom;
            layout3.Root.Padding = 0.25;
            layout3.Root.ChildSeparation = 0.25;
            layout3.Root.AddBox(1, 2);
            layout3.Root.AddBox(1, 1);

            // Now switch to horizontal containers
            var layout4 = new BoxL.BoxLayout();
            layout4.Root = new BoxL.Container(BoxL.ContainerDirection.Horizontal);
            layout4.Root.MinHeight = 5;
            layout4.Root.MinWidth = 3;
            layout4.Root.ChildHorizontalDirection = DirectionHorizontal.RightToLeft;
            layout4.Root.Padding = 0.25;
            layout4.Root.ChildSeparation = 0.25;
            layout4.Root.AddBox(1, 2);
            layout4.Root.AddBox(1, 1);


            Util.Render(layout1, doc);
            Util.Render(layout2, doc);
            Util.Render(layout3, doc);
            Util.Render(layout4, doc);


        }
    }

    public static class Util
    {
        public static void Render(BoxL.BoxLayout layout, IVisio.Document doc)
        {
            layout.PerformLayout();
            var page1 = doc.Pages.Add();
            // and tinker with it
            // render
            var nodes = layout.Nodes.ToList();
            foreach (var node in nodes)
            {
                var shape = page1.DrawRectangle(node.Rectangle);
                node.Data = shape;
            }

            var root_shape = (IVisio.Shape)layout.Root.Data;
            root_shape.CellsU["FillForegnd"].FormulaForceU = "rgb(240,240,240)";
            var margin = new VA.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);

        }
    
    }
}