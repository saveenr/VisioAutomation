using System;
using System.Collections.Generic;

namespace VisioScripting.Models
{
    public class TargetShapeIDs: List<int>
    {
        internal TargetShapeIDs(int capacity) : base (capacity)
        {
        }
    }
}