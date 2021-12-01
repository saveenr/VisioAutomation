using System.Collections.Generic;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyDictionary : Dictionary<string, CustomPropertyCells>
    {
        public CustomPropertyDictionary() : base()
        {
        }

        public CustomPropertyDictionary(int capacity) : base(capacity)
        {
        }

        internal static CustomPropertyDictionary FromPairs(List<CustomPropertyNameCellsPair> pairs)

        {
            var shape_custprop_dic = new CustomPropertyDictionary(pairs.Count);

            foreach (var pair in pairs)
            {
                shape_custprop_dic[pair.Name] = pair.Cells;
            }

            return shape_custprop_dic;
        }
    }
}