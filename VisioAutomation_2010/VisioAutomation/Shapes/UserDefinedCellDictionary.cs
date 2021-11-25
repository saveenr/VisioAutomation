using System.Collections.Generic;

namespace VisioAutomation.Shapes;

public class UserDefinedCellDictionary : Dictionary<string, UserDefinedCellCells>
{
    public UserDefinedCellDictionary(int capacity) : base(capacity)
    {

    }

    internal static UserDefinedCellDictionary FromPairs(List<UserDefinedCellNameCellsPair> pairs)

    {
        var dic = new UserDefinedCellDictionary(pairs.Count);
        foreach (var pair in pairs)
        {
            dic[pair.Name] = pair.Cells;
        }
        return dic;
    }

}