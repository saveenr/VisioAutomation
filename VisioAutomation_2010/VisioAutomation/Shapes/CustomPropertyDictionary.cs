using System.Collections.Generic;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyDictionary : Dictionary<string, CustomPropertyCells>
    {
        public CustomPropertyDictionary():base(){ }
        public CustomPropertyDictionary(int capacity):base(capacity){ }
    }
}