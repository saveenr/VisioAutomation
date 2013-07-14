using System.Collections;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Data
{

    /// <summary>
    /// Used to store the output of the QueryRows and QueryCells methods. Stores a string for the formula and a typed value (int|bool|double|string) for the result.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Table<T> 
    {
        private readonly T[] _values;
        private readonly int rowcount;
        private readonly int colcount;

        public List<VA.ShapeSheet.Data.TableRowGroup> Groups { get; private set; }

        internal Table(int rows, int cols, List<VA.ShapeSheet.Data.TableRowGroup> groups, T[] values)
        {
            int total_cells = rows*cols;
            if (values.Length != total_cells)
            {
                throw new VA.AutomationException("incorret number of values for rows and columns");
            }

            this._values = values;
            this.Groups = groups;
            this.rowcount = rows;
            this.colcount = cols;
        }

        private int get_pos(int row, int col)
        {
            if (row >= this.rowcount)
            {
                throw new System.ArgumentOutOfRangeException("row");
            }
            if (col >= this.colcount)
            {
                throw new System.ArgumentOutOfRangeException("col");
            }
            return (row * this.colcount) + col;
        }

        public T this[int row, int column]
        {
            get
            {
                return this._values[get_pos(row,column)];
            }
            set { this._values[get_pos(row, column)] = value; }
        }

        public T this[int row, VA.ShapeSheet.Query.CellQuery.Column column]
        {
            get
            {
                if (column==null)
                {
                    throw new System.ArgumentNullException("column");
                }
                return this._values[get_pos(row, column.Ordinal)];
            }
            set { this._values[get_pos(row, column.Ordinal)] = value; }
        }

        public int RowCount
        {
            get { return this.rowcount; }
        }

        internal T[] RawArray
        {
            get { return this._values; }
        }

        public int ColumnCount
        {
            get { return this.colcount; }
        }
    }
}