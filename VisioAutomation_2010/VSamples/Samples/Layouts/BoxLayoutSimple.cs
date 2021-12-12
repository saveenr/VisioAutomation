namespace VSamples.Samples.Layouts
{
    public class BoxLayoutSimple : SampleMethodBase
    {
        public override void RunSample()
        {
            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);

            // Create a simple Column
            var layout1 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout1.Root =
                new VisioAutomation.Models.Layouts.Box.Container(VisioAutomation.Models.Layouts.Box.Direction
                    .BottomToTop);
            layout1.Root.AddBox(1, 2);
            layout1.Root.AddBox(1, 1);
            layout1.Root.AddBox(0.5, 0.5);

            // You can set the min height and width of a container
            var layout2 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout2.Root =
                new VisioAutomation.Models.Layouts.Box.Container(
                    VisioAutomation.Models.Layouts.Box.Direction.BottomToTop, 3, 5);
            layout2.Root.AddBox(1, 2);
            layout2.Root.AddBox(1, 1);
            layout2.Root.AddBox(0.5, 0.5);

            // For vertical containers, you can layout shapes bottom-to-top or top-to-bottom
            var layout3 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout3.Root =
                new VisioAutomation.Models.Layouts.Box.Container(
                    VisioAutomation.Models.Layouts.Box.Direction.TopToBottom, 3, 5);
            layout3.Root.AddBox(1, 2);
            layout3.Root.AddBox(1, 1);
            layout3.Root.AddBox(0.5, 0.5);

            // Now switch to horizontal containers
            var layout4 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout4.Root =
                new VisioAutomation.Models.Layouts.Box.Container(
                    VisioAutomation.Models.Layouts.Box.Direction.RightToLeft, 3, 5);
            layout4.Root.AddBox(1, 2);
            layout4.Root.AddBox(1, 1);
            layout4.Root.AddBox(0.5, 0.5);


            // For Columns, you can tell the children how to horizontally align
            var layout5 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout5.Root =
                new VisioAutomation.Models.Layouts.Box.Container(
                    VisioAutomation.Models.Layouts.Box.Direction.BottomToTop, 3, 0);
            var b51 = layout5.Root.AddBox(1, 2);
            var b52 = layout5.Root.AddBox(1, 1);
            var b53 = layout5.Root.AddBox(0.5, 0.5);
            b51.HAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentHorizontal.Left;
            b52.HAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentHorizontal.Center;
            b53.HAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentHorizontal.Right;

            // For Rows , you can tell the children how to vertially align
            var layout6 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout6.Root =
                new VisioAutomation.Models.Layouts.Box.Container(
                    VisioAutomation.Models.Layouts.Box.Direction.LeftToRight, 0, 5);
            var b61 = layout6.Root.AddBox(1, 2);
            var b62 = layout6.Root.AddBox(1, 1);
            var b63 = layout6.Root.AddBox(0.5, 0.5);
            b61.VAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentVertical.Bottom;
            b62.VAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentVertical.Center;
            b63.VAlignToParent = VisioAutomation.Models.Layouts.Box.AlignmentVertical.Top;

            Util.BoxRender(layout1, doc);
            Util.BoxRender(layout2, doc);
            Util.BoxRender(layout3, doc);
            Util.BoxRender(layout4, doc);
            Util.BoxRender(layout5, doc);
            Util.BoxRender(layout6, doc);
        }
    }
}