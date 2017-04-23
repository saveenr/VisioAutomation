using System.Collections.Generic;

namespace VisioScripting.Models
{
    public class TargetShapeIDs
    {
        public readonly IList<int> ShapeIDs;

        public TargetShapeIDs(IList<int> shape_ids)
        {
            if (shape_ids == null)
            {
                throw new System.ArgumentNullException(nameof(shape_ids));
            }

            this.ShapeIDs = shape_ids;
        }
    }
}