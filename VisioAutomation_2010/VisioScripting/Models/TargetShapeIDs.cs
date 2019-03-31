using System;
using System.Collections.Generic;

namespace VisioScripting.Models
{
    public class TargetShapeIDs: List<int>
    {
        internal TargetShapeIDs(IEnumerable<int> shapeids, int count)
        {
            if (shapeids == null)
            {
                throw new System.ArgumentNullException(nameof(shapeids));
            }

            this.AddRange(shapeids);
        }
    }
}