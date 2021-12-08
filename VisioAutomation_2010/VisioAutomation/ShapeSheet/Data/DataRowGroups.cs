using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataRowGroups<T> : VisioAutomation.Core.BasicList<DataRowGroup<T>>
    {
        // Simple list of RowGroups

        internal DataRowGroups() : base()
        {
        }
    }
}