using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Results
{
    public class SubQueryResult<T>
    {
        public SubQuery Column { get; internal set; }
        public readonly List<SubQueryResultRow<T>> Rows;

        internal SubQueryResult(int capacity)
        {
            this.Rows = new List<SubQueryResultRow<T>>(capacity);
        }
    }
}