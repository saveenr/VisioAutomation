using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> ToEnumerable(this IVisio.Connects connects)
        {
            int count = connects.Count;
            for (int i = 0; i < count; i++)
            {
                yield return connects[i + 1];
            }
        }

        public static IList<IVisio.Connect> ToList(this IVisio.Connects connects)
        {
            int count = connects.Count;
            var list = new List<IVisio.Connect>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(connects[i+1]);
            }
            return list;
        }
    }
}
