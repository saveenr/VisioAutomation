using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class TextUtil
    {
        public static System.Text.RegularExpressions.Regex GetRegexForWildcardPattern(string wildcardpat, bool ignorecase)
        {
            string pat = "^" + System.Text.RegularExpressions.Regex.Escape(wildcardpat)
                .Replace(@"\*", ".*").
                Replace(@"\?", ".") + "$";

            var regexOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase;

            if (ignorecase)
            {
                var regex = new System.Text.RegularExpressions.Regex(pat, regexOptions);
                return regex;
            }
            else
            {
                var regex = new System.Text.RegularExpressions.Regex(pat);
                return regex;
            }
        }

        public static bool ContainsWildcard(string s)
        {
            return (s.Contains("*") || s.Contains("?"));
        }


        public static IEnumerable<string> IncludeByName(IEnumerable<string> items, IList<string> patterns, bool ignorecase)
        {
            return FilterObjectsByNames(items, patterns, System.IO.Path.GetFileName, ignorecase, FilterAction.Include);
        }

        public static IEnumerable<string> ExcludeByName(IEnumerable<string> items, IList<string> pattens, bool ignorecase)
        {
            return FilterObjectsByNames(items, pattens, System.IO.Path.GetFileName, ignorecase, FilterAction.Exclude);
        }

        public enum FilterAction
        {
            Include,
            Exclude
        }

        public static IEnumerable<T> FilterObjectsByNames<T>(IEnumerable<T> items, IList<string> patterns, Func<T, string> get_name, bool ignorecase, FilterAction action)
        {
            if (patterns == null || patterns.Count < 1)
            {
                // nothing to filter just return the items
                foreach (T item in items)
                {
                    yield return item;
                }
            }
            else
            {
                // Create the caches for fast matches of regexes
                var regexes = new List<System.Text.RegularExpressions.Regex>();

                HashSet<string> nonregexes;
                if (ignorecase)
                {
                    nonregexes = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                }
                else
                {
                    nonregexes = new HashSet<string>();
                }

                foreach (var pattern in patterns)
                {
                    if (TextUtil.ContainsWildcard(pattern))
                    {
                        // If it contains a wildcard transform it into a regex
                        var regex = TextUtil.GetRegexForWildcardPattern(pattern, ignorecase);
                        regexes.Add(regex);
                    }
                    else
                    {
                        // if it doesn't contain a wildcard then perform simple string equality
                        nonregexes.Add(pattern);
                    }
                }

                // the caches are set up, let's process each item

                foreach (var item in items)
                {
                    string name = get_name(item);

                    // does it match any of the patterns
                    // we test nonregexes first on the assumption that it's faster than checking regexes
                    bool matches = (nonregexes.Contains(name)) || (regexes.Any(regex => regex.IsMatch(name)));

                    if (action == FilterAction.Exclude && !matches)
                    {
                        // For exclude, non-match means this is a desired item so yield it 
                        yield return item;
                    }
                    else if (action == FilterAction.Include && matches)
                    {
                        // For include, match means this is a desired item so yield it 
                        yield return item;
                    }
                }
            }
        }
    }
}