using VisioAutomation.DOM;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BH = VisioAutomation.Layout.BoxHierarchy;

namespace VisioAutomationSamples
{
    public static class LayoutSamples
    {
        public class NodeData
        {
            public IVisio.Shape VisioShape;
            public string Text;
            public bool Render;
            public VA.DOM.ShapeCells Cells;
            public string Font;

            public NodeData()
            {
                this.Render = true;
                this.Cells = new VA.DOM.ShapeCells();
            }
        }

        public static BH.Node<NodeData> new_node()
        {
            return new_node(null);
        }

        public static BH.Node<NodeData> new_node(string s)
        {
            var box = new BH.Node<NodeData>();
            box.Data = new NodeData();
            box.Data.Text = s;
            return box;
        }

        public static BH.Node<NodeData> new_node(double w, double h, string s)
        {
            var box = new BH.Node<NodeData>();
            box.Width = w;
            box.Height = h;
            box.Data = new NodeData();
            box.Data.Text = s;
            return box;
        }




        public static void BoxHierarchy()
        {
            // Create a layout
            var layout = BoxHierarchyShared.CreateSampleBoxHierarchyLayout();

            // Ask the Layout to place the nodes
            var origin = new VA.Drawing.Point(0, 0);
            layout.LayoutOptions.Origin = origin;

            layout.PerformLayout();

            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);
            var page1 = doc.Pages[1];

            // and tinker with it
            // render
            foreach (var node in layout.Nodes)
            {
                BoxHierarchyShared.DrawBoxHierarchyDrawNode(node, node.Rectangle, page1);
            }

            var src_linepat = new VA.ShapeSheet.SRC(
                IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowLine,
                IVisio.VisCellIndices.visLinePattern);
            var root_shape = layout.Root.Data;
            var cell_linepat = root_shape.GetCell(src_linepat);
            cell_linepat.FormulaU = "7";

            // Make the page big enough to fit what was drawn + a small border
            var margin = new VA.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);
        }

        public static void BoxHeirarchy_FontGlyphComparision()
        {
            var sampletext = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "abcdefghijklmnopqrstuvwxyz" +
                             "<>[](),./|\\:;\'\"1234567890!@#$%^&*()`~";

            var samplechars = sampletext.Select(c => new string( new char[] { c })).ToList();

            var fontnames = new[] { "Calibri", "Arial" };

            var layout = new BH.BoxHierarchyLayout<NodeData>();
            layout.LayoutOptions.DirectionVertical = VA.DirectionVertical.TopToBottom;

            var root = layout.Root;
            root.Direction = BH.LayoutDirection.Vertical;
            root.ChildSeparation = 0.5;
            root.Data = new NodeData();
            root.Data.Render = false;

            var fontname_cells = new VA.DOM.ShapeCells();
            fontname_cells.FillPattern = 0;
            fontname_cells.LinePattern = 0;
            fontname_cells.LineWeight = 0.0;
            fontname_cells.HAlign = 0;
            fontname_cells.CharSize = VA.Convert.PointsToInches(36.0);

            var charbox_cells = new VA.DOM.ShapeCells();
            charbox_cells.FillPattern = 0;
            charbox_cells.LinePattern = 1;
            charbox_cells.LineWeight = 0.0;
            charbox_cells.LineColor= "rgb(150,150,150)";
            charbox_cells.HAlign = 1;
            charbox_cells.CharSize = VA.Convert.PointsToInches(24.0);

            foreach (string fontname in fontnames)
            {
                var fontname_box = new_node(5, 0.5, fontname);
                fontname_box.Data.Cells = fontname_cells;
                root.AddNode(fontname_box);

                var font_box = new_node();
                font_box.Direction = BH.LayoutDirection.Vertical;
                font_box.ChildSeparation = 0.25;
                font_box.Data.Render = false;
                root.AddNode(font_box);

                int numcols = 17;
                int numrows = 5;
                int numcells = numcols*numrows;


                foreach (int row in Enumerable.Range(0, numrows))
                {
                    var row_box = new_node();
                    row_box.Direction = BH.LayoutDirection.Horizonal;
                    row_box.ChildSeparation = 0.25;
                    row_box.Data.Render = false;
                    font_box.AddNode(row_box);

                    foreach (int col in Enumerable.Range(0, numcols))
                    {
                        int charindex = (col + (numcols*row))%numcells;
                        string curchar = samplechars[charindex];

                        var cell_box = new_node(0.50, 0.50, curchar);
                        cell_box.Data.Font = fontname;
                        cell_box.Data.Cells = charbox_cells;
                        row_box.AddNode(cell_box);
                    }
                }
            }

            layout.PerformLayout();

            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;
            var docs = visapp.Documents;
            var stencil = docs.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];

            
            var nodes = layout.Nodes.Where(n => n.Data.Render).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;

            var font_to_id = doc.Fonts.AsEnumerable().ToDictionary(f => f.Name, f => f.ID);

            foreach (var node in nodes)
            {
                var dom_shape = dom.Drop(rectmaster, node.Rectangle.Center);
                var cells = node.Data.Cells;
                if (cells == null)
                {
                    cells = new ShapeCells();
                }

                cells.Width = node.Rectangle.Width;
                cells.Height = node.Rectangle.Height;

                if (node.Data.Font != null)
                {
                    cells.CharFont = font_to_id[node.Data.Font];
                }
                cells.CharFont = 15;

                dom_shape.ShapeCells = cells;
                dom_shape.Text = node.Data.Text;
            }

            dom.Render(page);


            page.ResizeToFitContents(new VA.Drawing.Size(0.5, 0.5));

        }

        public static void MSAGL()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var vdom = new VisioAutomation.Layout.MSAGL.Drawing();

            // Create a Node 0
            var n0 = vdom.AddShape("n0", "N0 Untitled Node", "basflo_u.vss", "Decision");

            // Format Node 0
            n0.Size = new VA.Drawing.Size(3, 2);

            // Create Node 1

            var n1 = vdom.AddShape("n1", "N1", "basflo_u.vss", "Decision");

            // Format Node 1
            n1.ShapeCells = new VA.DOM.ShapeCells();
            n1.ShapeCells.FillForegnd = "rgb(255,0,0)";
            n1.ShapeCells.FillBkgnd = "rgb(255,255,0)";
            n1.ShapeCells.FillPattern = 40;

            // Create Node 2
            var n2 = vdom.AddShape("n2", "N2 MailServer", "server_u.vss", "Server");

            // Create Node 3

            var n3 = vdom.AddShape("n3", "N3", "basflo_u.vss", "Data");

            // Create Node 4
            var n4 = vdom.AddShape("n4", "N4", "basflo_u.vss", "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node
            var c0 = vdom.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = vdom.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c2 = vdom.Connect("c2", n3, n4, "NO", VA.Connections.ConnectorType.Curved);
            var c3 = vdom.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c4 = vdom.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c5 = vdom.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);

            // Format connector 0 to point "back" 
            c0.ShapeCells = new VA.DOM.ShapeCells();
            c0.ShapeCells.BeginArrow = 1;
            c0.ShapeCells.LineWeight = 0.10;

            // Format connector 1 to point "forward" 
            c1.ShapeCells = new VA.DOM.ShapeCells();
            c1.ShapeCells.EndArrow = 1;
            c1.ShapeCells.LineWeight = 0.10;

            // Format connector 2 to point "back" and "forward"  
            c2.ShapeCells = new VA.DOM.ShapeCells();
            c2.ShapeCells.EndArrow = 1;
            c2.ShapeCells.BeginArrow = 1;
            c2.ShapeCells.LineWeight = 0.10;

            // Perform the rendering
            var renderer = new VisioAutomation.Layout.MSAGL.DirectedGraphLayout();
            renderer.LayoutOptions.UseDynamicConnectors = true;

            var options = new VisioAutomation.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = false;

            vdom.Render(page1, options);
        }
    }

    
}