using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
namespace VisioAutomation.ShapeSheet.Data
{
    public struct QueryDataRow<T>
    {
        private QueryDataSet<T> _queryDataSet;
        private int _rowIndex;
        
        public QueryDataRow(QueryDataSet<T> qds, int row)
        {
            this._queryDataSet = qds;
            this._rowIndex = row;
        }

        public QueryDataSet<T> QueryDataSet
        {
            get { return _queryDataSet; }
        }

        public int RowIndex
        {
            get { return _rowIndex; }
        }

        public VA.ShapeSheet.CellData<T> this[VA.ShapeSheet.Query.QueryColumn col]
        {
            get
            {
                return this._queryDataSet.GetItem(this._rowIndex, col);
            }
        }
    }
}