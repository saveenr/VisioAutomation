using System.Collections.Generic;
using VACXN = VisioAutomation.Shapes.Connections;

namespace TestVisioAutomation
{
    public class ConnectivityMap
    {
        private Dictionary<string, List<string>> dic;

        public ConnectivityMap(IList<VACXN.ConnectorEdge> edges)
        {
            dic = new Dictionary<string, List<string>>();
            foreach (var e in edges)
            {
                string fromtext = e.From.Text;
                if (!dic.ContainsKey(fromtext))
                {
                    dic[fromtext] = new List<string>();
                }
                var list = dic[fromtext];
                list.Add(e.To.Text);
            }
        }

        public bool HasConnectionFromTo(string f, string t)
        {
            return (this.dic[f].Contains(t));
        }

        public int CountConnectionsFrom(string f)
        {
            return this.dic[f].Count;
        }

        public int CountFromNodes()
        {
            return this.dic.Count;
        }
    }
}