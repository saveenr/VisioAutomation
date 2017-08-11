using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VisioScripting.Models
{
    public class NamedDictionary<T>
    {
        private readonly Dictionary<string, T> dic;
        private readonly Regex regex_name;
        private readonly Regex regex_name_wildcard;

        public NamedDictionary()
        {
            this.regex_name = new Regex("^[a-zA-Z]*$");
            this.regex_name_wildcard = new Regex("^[a-zA-Z\\*\\?]*$");
            var compare = StringComparer.InvariantCultureIgnoreCase;
            this.dic = new Dictionary<string, T>(compare);
        }

        public List<string> GetNames()
        {
            return this.Keys.ToList();
        }

        public T this[string name]
        {
            get { return this.dic[name]; }
            set
            {
                this.CheckName(name);

                if (this.dic.ContainsKey(name))
                {
                    string msg = string.Format("Dictionary already contains a key called \"{0}\"", name);
                    throw new ArgumentOutOfRangeException(msg);
                }

                this.dic[name] = value;
            }
        }

        public Dictionary<string, T>.KeyCollection Keys
        {
            get { return this.dic.Keys; }
        }

        public bool IsValidName(string name)
        {
            return this.regex_name.IsMatch(name);
        }

        public bool IsValidNameWildCard(string name)
        {
            return this.regex_name_wildcard.IsMatch(name);
        }

        public void CheckName(string name)
        {
            if (this.IsValidName(name))
            {
                return;
            }

            string msg = string.Format("Key name \"{0}\" is not valid", name);
            throw new ArgumentOutOfRangeException(msg);
        }

        public void CheckNameWildcard(string name)
        {
            if (this.IsValidNameWildCard(name))
            {
                return;
            }

            string msg = string.Format("wildcard pattern \"{0}\" is not valid", name);
            throw new ArgumentException(msg, nameof(name));
        }

        public IEnumerable<string> ResolveName(string name)
        {
            if (name.Contains("*") || name.Contains("?"))
            {
                this.CheckNameWildcard(name);

                var regex = NamedDictionary<T>.GetRegexForWildCardPattern(name);

                foreach (string k in this.Keys)
                {
                    if (regex.IsMatch(k))
                    {
                        yield return k;
                    }
                }
            }
            else
            {
                this.CheckName(name);
                if (this.dic.ContainsKey(name))
                {
                    // found the exact cell name, yield it
                    yield return name;
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
            string pat = "^" + Regex.Escape(cellname).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";

            var regex = new Regex(pat, RegexOptions.IgnoreCase);
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

        public bool ContainsKey(string name)
        {
            return this.dic.ContainsKey(name);
        }
    }
}