using System.Collections.Generic;

namespace VisioScripting.Models
{
    public class TargetShapeIDs
    {
        public readonly IList<int> ShapeIDs;

        public TargetShapeIDs(IList<int> shapeids)
        {
            if (shapeids == null)
            {
                throw new System.ArgumentNullException(nameof(shapeids));
            }

            this.ShapeIDs = shapeids;
        }
    }
}