using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class GridBuilder
    {
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
        public  VA.Drawing.Size CellSize = new VA.Drawing.Size(2.0,1.5);
        
        public GridBuilder(int rows, int cols)
        {
            if (cols<1)
            {
                throw new System.ArgumentOutOfRangeException("cols");
            }

            if (rows<1)
            {
                throw new System.ArgumentOutOfRangeException("rows");
            }
            
            this.RowCount = rows;
            this.ColumnCount = cols;
        }

        public int Count
        {
            get { return this.RowCount*this.ColumnCount; }
        }

        public int GetRowsNeeded(int numitems)
        {
            int allocrows = System.Math.Max(1, (int)(numitems * 1.0 / this.ColumnCount + 0.5));
            return allocrows;
        }

        public VA.Drawing.Size Size
        {
            get
            {
                return this.CellSize.Multiply(this.ColumnCount, this.RowCount);
            }
        }

        public VA.Drawing.Rectangle GetCellRect(VA.Drawing.Point origin, int row, int col)
        {
            // Handle background cell
            var ul = origin.Add(col * this.CellSize.Width, -row * this.CellSize.Height);
            var cellrect = DocUtil.BuildFromUpperLeft(ul,this.CellSize);
            return cellrect;
        }

    }

}
