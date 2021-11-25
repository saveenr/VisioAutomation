using System.Collections.Generic;

namespace VisioAutomation_Tests
{
    public class ConnectivityMap
    {
        private readonly Dictionary<string, List<string>> _dic;

        public ConnectivityMap(IList<VisioAutomation.DocumentAnalysis.ConnectorEdge> edges)
        {
            this._dic = new Dictionary<string, List<string>>();
            foreach (var e in edges)
            {
                string fromtext = e.From.Text;
                if (!this._dic.ContainsKey(fromtext))
                {
                    this._dic[fromtext] = new List<string>();
                }
                var list = this._dic[fromtext];
                list.Add(e.To.Text);
            }
        }

        public bool HasConnectionFromTo(string f, string t)
        {
            return (this._dic[f].Contains(t));
        }

        public int CountConnectionsFrom(string f)
        {
            return this._dic[f].Count;
        }

        public int CountFromNodes()
        {
            return this._dic.Count;
        }
    }
}