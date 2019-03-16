using System.Collections.Generic;

namespace VisioAutomation.Shapes
{
    public class UserDefinedCellDictionary : Dictionary<string, UserDefinedCellCells>
    {
        public UserDefinedCellDictionary(int capacity) : base(capacity)
        {

        }

    }
}