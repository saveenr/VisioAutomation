using VAUDCELLS = VisioAutomation.Shapes.UserDefinedCells;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using VABOXMODEL = VisioAutomation.Models.BoxLayout;

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
            var layout1 = new VABOXMODEL.BoxLayout();
            layout1.Root = new VABOXMODEL.Container( VABOXMODEL.Direction.BottomToTop);
            layout1.Root.AddBox(1,2);
            layout1.Root.AddBox(1,1);
            layout1.Root.AddBox(0.5, 0.5);

            // You can set the min height and width of a container
            var layout2 = new VABOXMODEL.BoxLayout();
            layout2.Root = new VABOXMODEL.Container(VABOXMODEL.Direction.BottomToTop,3,5);
            layout2.Root.AddBox(1, 2);
            layout2.Root.AddBox(1, 1);
            layout2.Root.AddBox(0.5, 0.5);

            // For vertical containers, you can layout shapes bottom-to-top or top-to-bottom
            var layout3 = new VABOXMODEL.BoxLayout();
            layout3.Root = new VABOXMODEL.Container(VABOXMODEL.Direction.TopToBottom,3,5);
            layout3.Root.AddBox(1, 2);
            layout3.Root.AddBox(1, 1);
            layout3.Root.AddBox(0.5, 0.5);

            // Now switch to horizontal containers
            var layout4 = new VABOXMODEL.BoxLayout();
            layout4.Root = new VABOXMODEL.Container(VABOXMODEL.Direction.RightToLeft,3,5);
            layout4.Root.AddBox(1, 2);
            layout4.Root.AddBox(1, 1);
            layout4.Root.AddBox(0.5, 0.5);


            // For Columns, you can tell the children how to horizontally align
            var layout5 = new VABOXMODEL.BoxLayout();
            layout5.Root = new VABOXMODEL.Container(VABOXMODEL.Direction.BottomToTop,3,0);
            var b51 = layout5.Root.AddBox(1, 2);
            var b52 = layout5.Root.AddBox(1, 1);
            var b53 = layout5.Root.AddBox(0.5, 0.5);
            b51.HAlignToParent = VABOXMODEL.AlignmentHorizontal.Left;
            b52.HAlignToParent = VABOXMODEL.AlignmentHorizontal.Center;
            b53.HAlignToParent = VABOXMODEL.AlignmentHorizontal.Right;

            // For Rows , you can tell the children how to vertially align
            var layout6 = new VABOXMODEL.BoxLayout();
            layout6.Root = new VABOXMODEL.Container(VABOXMODEL.Direction.LeftToRight,0,5);
            var b61 = layout6.Root.AddBox(1, 2);
            var b62 = layout6.Root.AddBox(1, 1);
            var b63 = layout6.Root.AddBox(0.5, 0.5);
            b61.VAlignToParent = VABOXMODEL.AlignmentVertical.Bottom;
            b62.VAlignToParent = VABOXMODEL.AlignmentVertical.Center;
            b63.VAlignToParent = VABOXMODEL.AlignmentVertical.Top;

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
            public VA.DOM.ShapeCells ShapeCells;
        }

        public static void BoxLayout_TwoLevelGrouping()
        {
            int num_types = 10;
            int max_properties = 50;

            var types = typeof(VAUDCELLS.UserDefinedCell).Assembly.GetExportedTypes().Take(num_types).ToList();

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


            var domshapescol = new VA.DOM.ShapeList();
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
                    shape.Text = new VA.Text.Markup.TextElement(info.Text);                    
                }
                
                shape.Cells = info.ShapeCells.ShallowCopy();
            }
            domshapescol.Render(page);

            var bordersize = new VA.Drawing.Size(0.5, 0.5);
            page.ResizeToFitContents(bordersize);

        }

        private static VABOXMODEL.BoxLayout CreateTwoLevelLayout(List<string[]> data)
        {
            double itemsep = 0.0;
            var major_group_direction = VABOXMODEL.Direction.LeftToRight;
            var minor_group_direction = VABOXMODEL.Direction.TopToBottom;

            var name_to_major_group = new Dictionary<string, VABOXMODEL.Container>();
            var name_to_minor_group = new Dictionary<string, VABOXMODEL.Container>();

            var layout1 = new VABOXMODEL.BoxLayout();
            layout1.Root = new VABOXMODEL.Container(major_group_direction);

            var major_cells = new VA.DOM.ShapeCells();
            major_cells.FillForegnd = "rgb(245,245,245)";
            major_cells.CharFont = 0;
            major_cells.CharSize = "12pt";
            major_cells.ParaHorizontalAlign = "0";
            major_cells.VerticalAlign = "0";
            major_cells.LineWeight = "0";
            major_cells.LinePattern = "0";

            var minor_cells = new VA.DOM.ShapeCells();
            minor_cells.FillForegnd = "rgb(230,230,230)";
            minor_cells.CharFont = 0;
            minor_cells.CharSize = "10pt";
            minor_cells.ParaHorizontalAlign = "0";
            minor_cells.VerticalAlign = "0";
            minor_cells.LineWeight = "0";
            minor_cells.LinePattern = "0";

            var item_cells = new VA.DOM.ShapeCells();
            item_cells.CharFont = 0;
            item_cells.FillPattern = "0";
            item_cells.CharSize = "8pt";
            item_cells.ParaHorizontalAlign = "0";
            item_cells.VerticalAlign = "0";
            item_cells.LineWeight = "0";
            item_cells.LinePattern = "0";


            foreach (var row in data)
            {
                var majorname = row[0];
                var minorname = row[1];
                var itemname = row[2];

                VABOXMODEL.Container majorcnt;
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

                    VABOXMODEL.Box headerbox = majorcnt.AddBox(2, 0.25);
                }

                VABOXMODEL.Container minorcnt;
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

                    VABOXMODEL.Box headerbox = minorcnt.AddBox(2, 0.25);
                }

                VABOXMODEL.Box itembox = minorcnt.AddBox(2, 0.25);

                var item_info = new TwoLevelInfo();
                item_info.Text = itemname;
                item_info.Render = true;


                item_info.ShapeCells = item_cells;
                
                itembox.Data = item_info;
            }
            return layout1;
        }
    }

    public static class Util
    {
        public static void Render(VABOXMODEL.BoxLayout layout, IVisio.Document doc)
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