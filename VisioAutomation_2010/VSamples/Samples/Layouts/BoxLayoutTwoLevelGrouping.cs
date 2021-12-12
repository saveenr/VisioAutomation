using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;

namespace VSamples.Samples.Layouts
{
    public class BoxLayoutTwoLevelGrouping : SampleMethodBase
    {
        public override void RunSample()
        {
            int num_types = 10;
            int max_properties = 50;

            var types = typeof(VisioAutomation.Shapes.UserDefinedCellCells).Assembly.GetExportedTypes().Take(num_types)
                .ToList();

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

            var layout1 = BoxLayoutTwoLevelGrouping.CreateTwoLevelLayout(data);


            layout1.PerformLayout();

            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);
            var page = app.ActivePage;


            var domshapescol = new VisioAutomation.Models.Dom.ShapeList();
            //var rect_master = dom.m
            foreach (var item in layout1.Nodes)
            {
                if (item.Data == null)
                {
                    continue;
                }

                var info = (Util.BoxTwoLevelInfo) item.Data;

                if (!info.Render)
                {
                    continue;
                }

                var shape = domshapescol.Drop("Rectangle", "Basic_U.VSS", item.Rectangle);

                if (info.Text != null)
                {
                    shape.Text = new VisioAutomation.Models.Text.Element(info.Text);
                }

                shape.Cells = info.ShapeCells.ShallowCopy();
            }

            domshapescol.Render(page);

            var bordersize = new VisioAutomation.Core.Size(0.5, 0.5);
            page.ResizeToFitContents(bordersize);
        }

        private static VisioAutomation.Models.Layouts.Box.BoxLayout CreateTwoLevelLayout(List<string[]> data)
        {
            double itemsep = 0.0;
            var major_group_direction = VisioAutomation.Models.Layouts.Box.Direction.LeftToRight;
            var minor_group_direction = VisioAutomation.Models.Layouts.Box.Direction.TopToBottom;

            var name_to_major_group = new Dictionary<string, VisioAutomation.Models.Layouts.Box.Container>();
            var name_to_minor_group = new Dictionary<string, VisioAutomation.Models.Layouts.Box.Container>();

            var layout1 = new VisioAutomation.Models.Layouts.Box.BoxLayout();
            layout1.Root = new VisioAutomation.Models.Layouts.Box.Container(major_group_direction);

            var major_cells = new VisioAutomation.Models.Dom.ShapeCells();
            major_cells.FillForeground = "rgb(245,245,245)";
            major_cells.CharFont = 0;
            major_cells.CharSize = "12pt";
            major_cells.ParaHorizontalAlign = "0";
            major_cells.TextBlockVerticalAlign = "0";
            major_cells.LineWeight = "0";
            major_cells.LinePattern = "0";

            var minor_cells = new VisioAutomation.Models.Dom.ShapeCells();
            minor_cells.FillForeground = "rgb(230,230,230)";
            minor_cells.CharFont = 0;
            minor_cells.CharSize = "10pt";
            minor_cells.ParaHorizontalAlign = "0";
            minor_cells.TextBlockVerticalAlign = "0";
            minor_cells.LineWeight = "0";
            minor_cells.LinePattern = "0";

            var item_cells = new VisioAutomation.Models.Dom.ShapeCells();
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

                VisioAutomation.Models.Layouts.Box.Container majorcnt;
                if (name_to_major_group.ContainsKey(majorname))
                {
                    majorcnt = name_to_major_group[majorname];
                }
                else
                {
                    majorcnt = layout1.Root.AddContainer(minor_group_direction, 1, 1);

                    var major_info = new Util.BoxTwoLevelInfo();
                    major_info.Text = majorname;
                    major_info.Render = true;
                    major_info.ShapeCells = major_cells;
                    majorcnt.Data = major_info;


                    name_to_major_group[majorname] = majorcnt;

                    VisioAutomation.Models.Layouts.Box.Box headerbox = majorcnt.AddBox(2, 0.25);
                }

                VisioAutomation.Models.Layouts.Box.Container minorcnt;
                var minorkey = majorname + "___" + minorname;
                if (name_to_minor_group.ContainsKey(minorkey))
                {
                    minorcnt = name_to_minor_group[minorkey];
                }
                else
                {
                    minorcnt = majorcnt.AddContainer(minor_group_direction);
                    minorcnt.ChildSpacing = itemsep;
                    var minor_info = new Util.BoxTwoLevelInfo();
                    minor_info.Text = minorname;
                    minor_info.Render = true;
                    minor_info.ShapeCells = minor_cells;
                    minorcnt.Data = minor_info;
                    name_to_minor_group[minorkey] = minorcnt;

                    VisioAutomation.Models.Layouts.Box.Box headerbox = minorcnt.AddBox(2, 0.25);
                }

                VisioAutomation.Models.Layouts.Box.Box itembox = minorcnt.AddBox(2, 0.25);

                var item_info = new Util.BoxTwoLevelInfo();
                item_info.Text = itemname;
                item_info.Render = true;


                item_info.ShapeCells = item_cells;

                itembox.Data = item_info;
            }

            return layout1;
        }
    }
}