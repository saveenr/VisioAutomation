using VA = VisioAutomation;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using VisioAutomation.Models.Dom;
using VisioAutomation.Models.Layouts.Box;
using VisioAutomation.Shapes;

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
            var layout1 = new BoxLayout();
            layout1.Root = new Container( Direction.BottomToTop);
            layout1.Root.AddBox(1,2);
            layout1.Root.AddBox(1,1);
            layout1.Root.AddBox(0.5, 0.5);

            // You can set the min height and width of a container
            var layout2 = new BoxLayout();
            layout2.Root = new Container(Direction.BottomToTop,3,5);
            layout2.Root.AddBox(1, 2);
            layout2.Root.AddBox(1, 1);
            layout2.Root.AddBox(0.5, 0.5);

            // For vertical containers, you can layout shapes bottom-to-top or top-to-bottom
            var layout3 = new BoxLayout();
            layout3.Root = new Container(Direction.TopToBottom,3,5);
            layout3.Root.AddBox(1, 2);
            layout3.Root.AddBox(1, 1);
            layout3.Root.AddBox(0.5, 0.5);

            // Now switch to horizontal containers
            var layout4 = new BoxLayout();
            layout4.Root = new Container(Direction.RightToLeft,3,5);
            layout4.Root.AddBox(1, 2);
            layout4.Root.AddBox(1, 1);
            layout4.Root.AddBox(0.5, 0.5);


            // For Columns, you can tell the children how to horizontally align
            var layout5 = new BoxLayout();
            layout5.Root = new Container(Direction.BottomToTop,3,0);
            var b51 = layout5.Root.AddBox(1, 2);
            var b52 = layout5.Root.AddBox(1, 1);
            var b53 = layout5.Root.AddBox(0.5, 0.5);
            b51.HAlignToParent = AlignmentHorizontal.Left;
            b52.HAlignToParent = AlignmentHorizontal.Center;
            b53.HAlignToParent = AlignmentHorizontal.Right;

            // For Rows , you can tell the children how to vertially align
            var layout6 = new BoxLayout();
            layout6.Root = new Container(Direction.LeftToRight,0,5);
            var b61 = layout6.Root.AddBox(1, 2);
            var b62 = layout6.Root.AddBox(1, 1);
            var b63 = layout6.Root.AddBox(0.5, 0.5);
            b61.VAlignToParent = AlignmentVertical.Bottom;
            b62.VAlignToParent = AlignmentVertical.Center;
            b63.VAlignToParent = AlignmentVertical.Top;

            Util.Render(layout1, doc);
            Util.Render(layout2, doc);
            Util.Render(layout3, doc);
            Util.Render(layout4, doc);
            Util.Render(layout5, doc);
            Util.Render(layout6, doc);

        }

        public class TwoLevelInfo
        {
            public string Text;
            public bool Render;
            public ShapeCells ShapeCells;
        }

        public static void BoxLayout_TwoLevelGrouping()
        {
            int num_types = 10;
            int max_properties = 50;

            var types = typeof(UserDefinedCellCells).Assembly.GetExportedTypes().Take(num_types).ToList();

            var data = new List<string[]>();
            foreach (var type in types)
            {
                var properties = type.GetProperties().Take(max_properties).ToList();
                foreach (var property in properties)
                {
                    var item = new[] {type.Name, property.Name[0].ToString().ToUpper(), property.Name};
                    data.Add(item);
                }
            }

            var layout1 = BoxLayout2Samples.CreateTwoLevelLayout(data);


            layout1.PerformLayout();

            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);
            var page = app.ActivePage;


            var domshapescol = new ShapeList();
            //var rect_master = dom.m
            foreach (var item in layout1.Nodes)
            {
                if (item.Data ==null)
                {
                    continue;
                }
                var info = (TwoLevelInfo) item.Data;

                if (!info.Render)
                {
                    continue;
                }

                var shape = domshapescol.Drop("Rectangle", "Basic_U.VSS",item.Rectangle);

                if (info.Text!=null)
                {
                    shape.Text = new VisioAutomation.Models.Text.Element(info.Text);                    
                }
                
                shape.Cells = info.ShapeCells.ShallowCopy();
            }
            domshapescol.Render(page);

            var bordersize = new VA.Geometry.Size(0.5, 0.5);
            page.ResizeToFitContents(bordersize);

        }

        private static BoxLayout CreateTwoLevelLayout(List<string[]> data)
        {
            double itemsep = 0.0;
            var major_group_direction = Direction.LeftToRight;
            var minor_group_direction = Direction.TopToBottom;

            var name_to_major_group = new Dictionary<string, Container>();
            var name_to_minor_group = new Dictionary<string, Container>();

            var layout1 = new BoxLayout();
            layout1.Root = new Container(major_group_direction);

            var major_cells = new ShapeCells();
            major_cells.FillForeground = "rgb(245,245,245)";
            major_cells.CharFont = 0;
            major_cells.CharSize = "12pt";
            major_cells.ParaHorizontalAlign = "0";
            major_cells.TextBlockVerticalAlign = "0";
            major_cells.LineWeight = "0";
            major_cells.LinePattern = "0";

            var minor_cells = new ShapeCells();
            minor_cells.FillForeground = "rgb(230,230,230)";
            minor_cells.CharFont = 0;
            minor_cells.CharSize = "10pt";
            minor_cells.ParaHorizontalAlign = "0";
            minor_cells.TextBlockVerticalAlign = "0";
            minor_cells.LineWeight = "0";
            minor_cells.LinePattern = "0";

            var item_cells = new ShapeCells();
            item_cells.CharFont = 0;
            item_cells.FillPattern = "0";
            item_cells.CharSize = "8pt";
            item_cells.ParaHorizontalAlign = "0";
            item_cells.TextBlockVerticalAlign = "0";
            item_cells.LineWeight = "0";
            item_cells.LinePattern = "0";


            foreach (var row in data)
            {
                var majorname = row[0];
                var minorname = row[1];
                var itemname = row[2];

                Container majorcnt;
                if (name_to_major_group.ContainsKey(majorname))
                {
                    majorcnt = name_to_major_group[majorname];
                }
                else
                {
                    majorcnt = layout1.Root.AddContainer(minor_group_direction, 1, 1);

                    var major_info = new TwoLevelInfo();
                    major_info.Text = majorname;
                    major_info.Render = true;
                    major_info.ShapeCells = major_cells;
                    majorcnt.Data = major_info;
                    

                    name_to_major_group[majorname] = majorcnt;

                    Box headerbox = majorcnt.AddBox(2, 0.25);
                }

                Container minorcnt;
                var minorkey = majorname + "___" + minorname;
                if (name_to_minor_group.ContainsKey(minorkey))
                {
                    minorcnt = name_to_minor_group[minorkey];
                }
                else
                {
                    minorcnt = majorcnt.AddContainer(minor_group_direction);
                    minorcnt.ChildSpacing = itemsep;
                    var minor_info = new TwoLevelInfo();
                    minor_info.Text = minorname;
                    minor_info.Render = true;
                    minor_info.ShapeCells = minor_cells;
                    minorcnt.Data = minor_info;
                    name_to_minor_group[minorkey] = minorcnt;

                    Box headerbox = minorcnt.AddBox(2, 0.25);
                }

                Box itembox = minorcnt.AddBox(2, 0.25);

                var item_info = new TwoLevelInfo();
                item_info.Text = itemname;
                item_info.Render = true;


                item_info.ShapeCells = item_cells;
                
                itembox.Data = item_info;
            }
            return layout1;
        }
    }
}