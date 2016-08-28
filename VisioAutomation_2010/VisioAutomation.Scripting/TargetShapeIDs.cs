using System.Collections.Generic;

namespace VisioAutomation.Scripting
{
    public class TargetShapeIDs
    {
        public readonly IList<int> ShapeIDs;
        public readonly Microsoft.Office.Interop.Visio.Page Page;

        public TargetShapeIDs(Microsoft.Office.Interop.Visio.Page page, IList<int> shape_ids)
        {
            this.Page = page;
            this.ShapeIDs = shape_ids;
        }
    }
}