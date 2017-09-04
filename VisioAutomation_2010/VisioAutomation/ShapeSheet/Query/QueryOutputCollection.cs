using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryOutputCollectionCells<T> : List<QueryOutputCells<T>>
    {
        internal QueryOutputCollectionCells() : base()
        {
        }
    }

    public class QueryOutputCollectionSections<T> : List<QueryOutputSections<T>>
    {
        internal QueryOutputCollectionSections() : base()
        {
        }
    }
}