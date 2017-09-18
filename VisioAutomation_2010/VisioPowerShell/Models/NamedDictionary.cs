using System;
using System.Collections.Generic;

namespace VisioPowerShell.Models
{
    public class NamedDictionary<T>
    {
        private readonly Dictionary<string, T> dic;
        private readonly System.Text.RegularExpressions.Regex regex_name;
        private readonly System.Text.RegularExpressions.Regex regex_name_wildcard;

        public NamedDictionary()
        {
            this.regex_name = new System.Text.RegularExpressions.Regex("^[a-zA-Z]*$");
            this.regex_name_wildcard = new System.Text.RegularExpressions.Regex("^[a-zA-Z\\*\\?]*$");
            var compare = StringComparer.InvariantCultureIgnoreCase;
            this.dic = new Dictionary<string, T>(compare);
        }

        public T this[string name]
        {
            get { return this.dic[name]; }
            set
            {
                this._AssertKeyIsValid(name);

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

        private bool _IsValidKey(string name)
        {
            return this.regex_name.IsMatch(name);
        }

        private bool _IsValidKeyWithWildCard(string name)
        {
            return this.regex_name_wildcard.IsMatch(name);
        }

        private void _AssertKeyIsValid(string name)
        {
            if (this._IsValidKey(name))
            {
                return;
            }

            string msg = string.Format("Key name \"{0}\" is not valid", name);
            throw new ArgumentOutOfRangeException(msg);
        }

        private void _CheckNameWildcard(string name)
        {
            if (this._IsValidKeyWithWildCard(name))
            {
                return;
            }

            string msg = string.Format("wildcard pattern \"{0}\" is not valid", name);
            throw new ArgumentException(msg, nameof(name));
        }

        public IEnumerable<string> ExpandKeyWildcard(string key)
        {
            string str_asterisk = "*";
            string str_questionmark = "?";

            if (key.Contains(str_asterisk) || key.Contains(str_questionmark))
            {
                this._CheckNameWildcard(key);

                var regex = NamedDictionary<T>.GetRegexForWildCardPattern(key);

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
                this._AssertKeyIsValid(key);
                if (this.dic.ContainsKey(key))
                {
                    // found the exact cell name, yield it
                    yield return key;
                }
                else
                {
                    // Coudn't find the exact cell name, yield nothing
                    yield break;
                }
            }
        }

        private static System.Text.RegularExpressions.Regex GetRegexForWildCardPattern(string s)
        {
            string pat = "^" + System.Text.RegularExpressions.Regex.Escape(s).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
            var regex_options = System.Text.RegularExpressions.RegexOptions.IgnoreCase;
            var regex = new System.Text.RegularExpressions.Regex(pat, regex_options);
            return regex;
        }

        public bool ContainsKey(string key)
        {
            return this.dic.ContainsKey(key);
        }
    }
}