using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VisioPowerShell
{
    public class CellNameDictionary<T>
    {
        private readonly Dictionary<string, T> dic;
        private readonly Regex regex_cellname;
        private readonly Regex regex_cellname_wildcard;

        public CellNameDictionary()
        {
            this.regex_cellname = new Regex("^[a-zA-Z]*$");
            this.regex_cellname_wildcard = new Regex("^[a-zA-Z\\*\\?]*$");
            var compare = StringComparer.InvariantCultureIgnoreCase;
            this.dic = new Dictionary<string, T>(compare);
        }

        public List<string> GetNames()
        {
            return this.CellNames.ToList();
        }

        public T this[string name]
        {
            get { return this.dic[name]; }
            set
            {
                this.CheckCellName(name);

                if (this.dic.ContainsKey(name))
                {
                    string msg = $"CellMap already contains a cell called \"{name}\"";
                    throw new ArgumentOutOfRangeException(msg);
                }

                this.dic[name] = value;
            }
        }

        public Dictionary<string, T>.KeyCollection CellNames => this.dic.Keys;

        public bool IsValidCellName(string name) => this.regex_cellname.IsMatch(name);

        public bool IsValidCellNameWildCard(string name) => this.regex_cellname_wildcard.IsMatch(name);


        public void CheckCellName(string name)
        {
            if (this.IsValidCellName(name))
            {
                return;
            }

            string msg = $"Cell name \"{name}\" is not valid";
            throw new ArgumentOutOfRangeException(msg);
        }

        public void CheckCellNameWildcard(string name)
        {
            if (this.IsValidCellNameWildCard(name))
            {
                return;
            }

            string msg = $"Cell name wildcard pattern \"{name}\" is not valid";
            throw new ArgumentException(msg, nameof(name));
        }

        public IEnumerable<string> ResolveName(string cellname)
        {
            if (cellname.Contains("*") || cellname.Contains("?"))
            {
                this.CheckCellNameWildcard(cellname);

                var regex = CellNameDictionary<T>.GetRegexForWildCardPattern(cellname);

                foreach (string k in this.CellNames)
                {
                    if (regex.IsMatch(k))
                    {
                        yield return k;
                    }
                }
            }
            else
            {
                this.CheckCellName(cellname);
                if (this.dic.ContainsKey(cellname))
                {
                    // found the exact cell name, yield it
                    yield return cellname;
                }
                else
                {
                    // Coudn't find the exact cell name, yield nothing
                    yield break;
                }
            }
        }

        private static Regex GetRegexForWildCardPattern(string cellname)
        {
            string pat = "^" + Regex.Escape(cellname)
                .Replace(@"\*", ".*").
                Replace(@"\?", ".") + "$";

            var regex = new Regex(pat,
                RegexOptions.IgnoreCase);
            return regex;
        }

        public IEnumerable<string> ResolveNames(IEnumerable<string> cellnames)
        {
            foreach (var name in cellnames)
            {
                foreach (var resolved_name in this.ResolveName(name))
                {
                    yield return resolved_name;
                }
            }
        }

        public bool ContainsCell(string name) => this.dic.ContainsKey(name);
    }
}