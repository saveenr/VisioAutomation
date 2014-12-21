using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Grid
{
    public class GridLayout
    {
        public int ColumnCount { get; private set; }
        public int RowCount { get; private set; }
        public VA.Drawing.Point Origin { get;  set; }
        public VA.Drawing.Size CellSpacing { get; set; }
        public RowDirection RowDirection { get; set; }
        public ColumnDirection ColumnDirection { get; set; }

        public IList<Column> Columns { get; private set; }
        public IList<Row> Rows { get; private set; }

        private readonly Node[,] _nodes;

        public IEnumerable<Node> Nodes
        {
            get
            {
                foreach (int row in Enumerable.Range(0, this.RowCount))
                {
                    foreach (int col in Enumerable.Range(0, this.ColumnCount))
                    {
                        var node = this._nodes[row, col];
                        yield return node;
                    }
                }
            }
        }

        public GridLayout(int cols, int rows, VA.Drawing.Size cellsize, IVisio.Master master)
        {
            ColumnDirection = ColumnDirection.LeftToRight;
            RowDirection = RowDirection.BottomToTop;
            CellSpacing = new VA.Drawing.Size(0.5, 0.25);
            this.ColumnCount = cols;
            this.RowCount = rows;

            // initialize the sizes for the rows and columns
            this.Rows = new List<Row>(this.RowCount);
            foreach (int row in Enumerable.Range(0, this.RowCount))
            {
                var r = new Row();
                r.Height = cellsize.Height;
                this.Rows.Add(r);
            }

            this.Columns = new List<Column>(this.ColumnCount);
            foreach (int col in Enumerable.Range(0, this.ColumnCount))
            {
                var c = new Column();
                c.Width = cellsize.Width;
                this.Columns.Add(c);
            }

            // Create the nodes
            this._nodes = new Node[this.RowCount, this.ColumnCount];
            foreach (int row in Enumerable.Range(0, this.RowCount))
            {
                foreach (int col in Enumerable.Range(0, this.ColumnCount))
                {
                    var node = new Node();
                    node.Column = col;
                    node.Row = row;
                    node.Master = master;
                    node.Draw = true;
                    this._nodes[row, col] = node;
                }
            }
        }

        public Node GetNode(int col, int row)
        {
            return this._nodes[row, col];
        }

        public void PerformLayout()
        {
            double dy = 0.0;

            foreach (int row in Enumerable.Range(0, this.RowCount))
            {
                // Restart calculating the cols
                double dx = 0;
                foreach (int col in Enumerable.Range(0, this.ColumnCount))
                {
                    double final_left;
                    double final_right;
                    double final_top;
                    double final_bottom;

                    if (ColumnDirection == ColumnDirection.LeftToRight)
                    {
                        final_left = this.Origin.X + dx;
                        final_right = final_left + this.Columns[col].Width;                       
                    }
                    else
                    {
                        final_right = this.Origin.X - dx;
                        final_left = final_right - this.Columns[col].Width;
                    }

                    if (RowDirection==RowDirection.BottomToTop)
                    {
                        final_bottom = this.Origin.Y + dy;
                        final_top = final_bottom + this.Rows[row].Height;
                    }
                    else
                    {
                        final_top = this.Origin.Y - dy;
                        final_bottom = final_top - this.Rows[row].Height;        
                    }

                    var node = this.GetNode(col, row);
                    node.Rectangle = new VA.Drawing.Rectangle(final_left, final_bottom, final_right, final_top);

                    dx += this.Columns[col].Width;
                    dx += this.CellSpacing.Width;
                }

                dy += this.Rows[row].Height;
                dy += this.CellSpacing.Height;
            }
        }

        public void Render(IVisio.Page page)
        {
            if (page == null)
            {
                throw new ArgumentNullException("page");
            }

            var nodes_to_draw = this.Nodes.Where(n => n.Draw).ToList();

            var page_node = new VA.DOM.Page();

            var shape_nodes = new List<VA.DOM.Shape>(nodes_to_draw.Count);
            foreach (var node in nodes_to_draw)
            {
                var shape_node = page_node.Shapes.Drop(node.Master, node.Rectangle.Center);

                if (node.Cells != null)
                {
                    shape_node.Cells = node.Cells.ShallowCopy();
                }

                shape_node.Cells.Width = node.Rectangle.Width;
                shape_node.Cells.Height = node.Rectangle.Height;

                if (!string.IsNullOrEmpty(node.Text))
                {
                    shape_node.Text = new VA.Text.Markup.TextElement( node.Text );
                }

                shape_nodes.Add(shape_node);
            }

            page_node.Shapes.Render(page);

            for (int i = 0; i < nodes_to_draw.Count; i++)
            {
                var node = nodes_to_draw[i];
                var shape_node = shape_nodes[i];

                node.Shape = shape_node.VisioShape;
                node.ShapeID = shape_node.VisioShapeID;
            }
        }
    }
}