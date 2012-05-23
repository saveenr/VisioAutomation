using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BOXMODEL = VisioAutomation.Layout.Models.BoxLayout;

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
            var layout1 = new BOXMODEL.BoxLayout();
            layout1.Root = new BOXMODEL.Container( BOXMODEL.Direction.BottomToTop);
            layout1.Root.AddBox(1,2);
            layout1.Root.AddBox(1,1);
            layout1.Root.AddBox(0.5, 0.5);

            // You can set the min height and width of a container
            var layout2 = new BOXMODEL.BoxLayout();
            layout2.Root = new BOXMODEL.Container(BOXMODEL.Direction.BottomToTop,3,5);
            layout2.Root.AddBox(1, 2);
            layout2.Root.AddBox(1, 1);
            layout2.Root.AddBox(0.5, 0.5);

            // For vertical containers, you can layout shapes bottom-to-top or top-to-bottom
            var layout3 = new BOXMODEL.BoxLayout();
            layout3.Root = new BOXMODEL.Container(BOXMODEL.Direction.TopToBottom,3,5);
            layout3.Root.AddBox(1, 2);
            layout3.Root.AddBox(1, 1);
            layout3.Root.AddBox(0.5, 0.5);

            // Now switch to horizontal containers
            var layout4 = new BOXMODEL.BoxLayout();
            layout4.Root = new BOXMODEL.Container(BOXMODEL.Direction.RightToLeft,3,5);
            layout4.Root.AddBox(1, 2);
            layout4.Root.AddBox(1, 1);
            layout4.Root.AddBox(0.5, 0.5);


            // For Columns, you can tell the children how to horizontally align
            var layout5 = new BOXMODEL.BoxLayout();
            layout5.Root = new BOXMODEL.Container(BOXMODEL.Direction.BottomToTop,3,0);
            var b51 = layout5.Root.AddBox(1, 2);
            var b52 = layout5.Root.AddBox(1, 1);
            var b53 = layout5.Root.AddBox(0.5, 0.5);
            b51.HAlignToParent = BOXMODEL.AlignmentHorizontal.Left;
            b52.HAlignToParent = BOXMODEL.AlignmentHorizontal.Center;
            b53.HAlignToParent = BOXMODEL.AlignmentHorizontal.Right;

            // For Rows , you can tell the children how to vertially align
            var layout6 = new BOXMODEL.BoxLayout();
            layout6.Root = new BOXMODEL.Container(BOXMODEL.Direction.LeftToRight,0,5);
            var b61 = layout6.Root.AddBox(1, 2);
            var b62 = layout6.Root.AddBox(1, 1);
            var b63 = layout6.Root.AddBox(0.5, 0.5);
            b61.VAlignToParent = BOXMODEL.AlignmentVertical.Bottom;
            b62.VAlignToParent = BOXMODEL.AlignmentVertical.Center;
            b63.VAlignToParent = BOXMODEL.AlignmentVertical.Top;

            Util.Render(layout1, doc);
            Util.Render(layout2, doc);
            Util.Render(layout3, doc);
            Util.Render(layout4, doc);
            Util.Render(layout5, doc);
            Util.Render(layout6, doc);

        }
    }

    public static class Util
    {
        public static void Render(BOXMODEL.BoxLayout layout, IVisio.Document doc)
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