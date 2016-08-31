using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class TargetShapeIDs
    {
        public readonly IList<int> ShapeIDs;
        public readonly IVisio.Page Page;

        public TargetShapeIDs(IVisio.Page page, IList<int> shape_ids)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            if (shape_ids == null)
            {
                throw new System.ArgumentNullException(nameof(shape_ids));
            }

            this.Page = page;
            this.ShapeIDs = shape_ids;
        }
    }
}