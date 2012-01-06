using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.DOM;
using VisioAutomation.Drawing;
using VisioAutomation.Layout.BoxLayout;
using BL = VisioAutomation.Layout.BoxLayout;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    class RenderItem
    {
        public CategoryCell CategoryCell;
        public string ShapeText ;
        public VA.DOM.ShapeCells ShapeCells;
        public bool Underline;
        public bool FitWidthToParent;
    }

    public class CategoryChart
    {
        public List<CategoryCell> Items;

        public string Font="Segoe UI";
        public bool ToUpper;
        public string Title = "Untitled";
        double TitleFontSize = 24;
        double CellFontSize = 8;
        double CategoryFontSize = 14;

        double CellWidth = 3.0;
        double CellVerticalSeparation = 0.125;
        double CellHeight = 0.25;
        double Indent = 2.0;
        double CategoryHeight = 0.5;
        double CellHorizontalSeparation = 0.25;

        public string CellFill = "rgb(240,240,240)";
        public string SubCellFill = "rgb(220,220,220)";

        ShapeCells titleformat = new VA.DOM.ShapeCells();
        ShapeCells cellformat = new VA.DOM.ShapeCells();
        ShapeCells subcellformat = new VA.DOM.ShapeCells();
        ShapeCells xcatformat = new VA.DOM.ShapeCells();
        ShapeCells ycatformat = new VA.DOM.ShapeCells();

        public CategoryChart()
        {
            this.Items = new List<CategoryCell>();

            titleformat.VerticalAlign = 0;
            titleformat.HAlign = 0;
            titleformat.CharSize = VA.Convert.PointsToInches(TitleFontSize);
            titleformat.LinePattern = 0;
            titleformat.LineWeight = 0;

            cellformat.VerticalAlign = 0;
            cellformat.HAlign = 0;
            cellformat.CharSize = VA.Convert.PointsToInches(CellFontSize);
            cellformat.LinePattern = 0;
            cellformat.LineWeight = 0;
            cellformat.FillForegnd = CellFill;

            subcellformat.VerticalAlign = 0;
            subcellformat.HAlign = 0;
            subcellformat.CharSize = VA.Convert.PointsToInches(CellFontSize);
            subcellformat.LinePattern = 0;
            subcellformat.LineWeight = 0;
            subcellformat.FillForegnd = SubCellFill;

            xcatformat.VerticalAlign = 2;
            xcatformat.HAlign = 1;
            xcatformat.CharSize = VA.Convert.PointsToInches(CategoryFontSize);
            xcatformat.LinePattern = 0;
            xcatformat.LineWeight = 0;
            xcatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

            ycatformat.VerticalAlign = 2;
            ycatformat.HAlign = 0;
            ycatformat.CharSize = VA.Convert.PointsToInches(CategoryFontSize);
            ycatformat.LinePattern = 0;
            ycatformat.LineWeight = 0;
            ycatformat.CharStyle = ((int)VA.Text.CharStyle.Bold);

        }

        public CategoryCell Add(string text, string xcat, string ycat)
        {
            var item = new CategoryCell(text,xcat,ycat);
            this.Items.Add(item);
            return item;
        }

        public CategoryCell Add(string text, string xcat, string ycat, IList<string> subitems)
        {
            var item = new CategoryCell(text, xcat, ycat);

            item.Item.Items = subitems.Select(t=>new CategoryItem(t)).ToList();
            this.Items.Add(item);
            return item;
        }

        public void Render(IVisio.Page page)
        {
            var xcats = this.Items.Select(i => i.XCategory).Distinct().ToList();
            var ycats = this.Items.Select(i => i.YCategory).Distinct().ToList();

            int cols = xcats.Count();
            int rows = ycats.Count();

            VA.Layout.BoxLayout.Node root;
            var layout = create_layout(out root);

            foreach (int row in Enumerable.Range(0, rows))
            {
                AddMajorRow(ycats, row, root, xcats, cols);
            }

            AddXCatLabels(xcats, cols, root);

            // Add Title for Chart
            add_title(root);

            Render(page, layout);
        }

        private void AddXCatLabels(List<string> xcats, int cols, VA.Layout.BoxLayout.Node root)
        {
            var n_row = root.AddRow();
            n_row.ChildSeparation = CellHorizontalSeparation;

            // Add indent
            n_row.AddBox(Indent, 0.25);

            // Add XCategory labels
            foreach (int col in Enumerable.Range(0, cols))
            {
                var n_label = n_row.AddBox(CellWidth, 0.5);
                var info = new RenderItem();
                info.CategoryCell = null;
                info.ShapeText = xcats[col];
                info.ShapeCells = xcatformat;
                n_label.Data = info;
            }
        }

        private void AddMajorRow(List<string> ycats, int row, VA.Layout.BoxLayout.Node root, List<string> xcats, int cols)
        {
            var n_row = root.AddRow();
            n_row.ChildSeparation = CellHorizontalSeparation;

            // -- add indent
            n_row.AddBox(Indent, 0.25);

            foreach (int col in Enumerable.Range(0, cols))
            {
                var n_cell = n_row.AddBox(CellWidth, 0.25);

                // ---
                n_cell.Direction = BL.LayoutDirection.Vertical;
                n_cell.AlignmentVertical = AlignmentVertical.Top;
                n_cell.ChildSeparation = CellVerticalSeparation;
                var items_for_cells = this.Items.Where(i => i.XCategory == xcats[col] && i.YCategory == ycats[row]);
                foreach (var cell_item in items_for_cells)
                {
                    draw_cell(cell_item, n_cell);
                }
            }

            var n_row_label = root.AddBox(null, CategoryHeight);
            var info = new RenderItem();
            info.CategoryCell = null;
            info.ShapeText = ycats[row];
            info.ShapeCells = ycatformat;
            info.Underline = true;
            info.FitWidthToParent = true;
            n_row_label.Data = info;
        }

        private BoxLayout create_layout(out VA.Layout.BoxLayout.Node root)
        {
            var layout = new BL.BoxLayout();
            layout.LayoutOptions.Origin = new VA.Drawing.Point(0, 10);
            layout.LayoutOptions.DefaultHeight = 0.25;
            root = layout.Root;
            root.Direction = BL.LayoutDirection.Vertical;
            root.ChildSeparation = 0.125;
            return layout;
        }

        private void add_title(VA.Layout.BoxLayout.Node root)
        {
            var n_title = root.AddBox(2.0, 0.5);
            var node_data = new RenderItem();
            node_data.CategoryCell = null;
            node_data.ShapeText = this.Title;
            node_data.ShapeCells = titleformat;
            node_data.FitWidthToParent = true;
            n_title.Data = node_data;
        }

        private void draw_cell(CategoryCell cell_item, VA.Layout.BoxLayout.Node n_row_col)
        {
            var n_cell = n_row_col.AddBox(CellWidth, CellHeight);
            n_cell.ChildSeparation = CellVerticalSeparation/2;
            
            var cell_data = new RenderItem();
            cell_data.CategoryCell = cell_item;
            cell_data.ShapeText = cell_item.Item.Text;
            cell_data.ShapeCells = cellformat;
            n_cell.Data = cell_data;
            
            if (cell_item.Item.Items != null)
            {
                foreach (var sub_cat_items in cell_item.Item.Items)
                {
                    var subn_cell = n_cell.AddBox(CellWidth, CellHeight);
                    var subcell_data = new RenderItem();
                    subcell_data.CategoryCell = null;
                    subcell_data.ShapeText = sub_cat_items.Text;
                    subcell_data.ShapeCells = subcellformat;
                    subn_cell.Data = subcell_data;
                }
                n_cell.AddBox(null, 0.25);
            }
        }

        private void Render(Page page, BoxLayout layout)
        {
            layout.PerformLayout();
            var doc = page.Document;
            var fonts = doc.Fonts;
            var default_font = fonts[this.Font];
            int default_font_id = default_font.ID;
            // Perform Rendering
            var dom = new VA.DOM.Document();
            foreach (var n in layout.Nodes)
            {
                if (n.Data != null)
                {
                    var r = n.Rectangle;
                    var n_data = (RenderItem) n.Data;
                    if (n_data.FitWidthToParent == true)
                    {
                       r = new VA.Drawing.Rectangle(r.LowerLeft, new VA.Drawing.Size(n.Parent.Width.Value-2*n.Padding,r.Height));
                    }

                    var s = dom.DrawRectangle(r);

                    // Set Text
                    if (n_data.ShapeText != null)
                    {
                        s.TextElement = new VA.Text.Markup.TextElement(this.ToUpper ? n_data.ShapeText.ToUpper() : n_data.ShapeText);
                    }

                    // Set Cells
                    if (n_data.ShapeCells != null)
                    {
                        s.ShapeCells = n_data.ShapeCells;
                    }

                    // draw Underline
                    if (n_data.Underline)
                    {
                        var u = dom.DrawLine(r.LowerLeft, r.LowerRight);
                    }

                    n_data.ShapeCells.CharFont = default_font_id;
                }
            }
            dom.Render(page);
        }

        public static CategoryChart FromCSV(string title, string text)
        {
            var chart = new CategoryChart();
            chart.Title = title;
            foreach (var line in text.Split(new char[] { '\n' }))
            {
                var sline = line.Trim();
                if (sline.Length < 1)
                {
                    continue;
                }

                var tokens = line.Split(new char[] { ',' });
                if (tokens.Length < 3)
                {
                    throw new System.Exception("Not enough tokens in line");
                }

                string xcat = tokens[0];
                string ycat = tokens[1];
                string item = tokens[2];
                string[] subitems = tokens.Length >= 4
                                        ? tokens[3].Split(new char[] { '|' }).Select(s => s.Trim()).Where(s => s.Length > 0).
                                              ToArray()
                                        : null;
                if (subitems == null)
                {
                    chart.Add(item, xcat, ycat);
                }
                else
                {
                    chart.Add(item, xcat, ycat, subitems);
                }
            }

            return chart;
        }

    }
}
