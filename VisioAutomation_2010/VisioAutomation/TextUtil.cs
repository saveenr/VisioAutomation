using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class TextUtil
    {
        public static System.Text.RegularExpressions.Regex GetRegexForWildcardPattern(string wildcardpat,
            bool ignorecase)
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


        public static IEnumerable<string> IncludeByName(IEnumerable<string> filenames, string[] exclude_patterns)
        {
            return FilterObjectsByNames(filenames, exclude_patterns, System.IO.Path.GetFileName, FilterAction.Include);
        }

        public static IEnumerable<string> ExcludeByName(IEnumerable<string> filenames, string[] exclude_patterns)
        {
            return FilterObjectsByNames(filenames, exclude_patterns, System.IO.Path.GetFileName, FilterAction.Exclude);
        }

        public enum FilterAction
        {
            Include,
            Exclude
        }

        public static IEnumerable<T> FilterObjectsByNames<T>(IEnumerable<T> items, IList<string> patterns,
            System.Func<T, string> get_name, FilterAction action)
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
                var regexes = new List<System.Text.RegularExpressions.Regex>();
                var nonregexes = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

                foreach (var pattern in patterns)
                {
                    if (TextUtil.ContainsWildcard(pattern))
                    {
                        // If it contains a wildcard transform it into a regex
                        var regex = TextUtil.GetRegexForWildcardPattern(pattern, true);
                        regexes.Add(regex);
                    }
                    else
                    {
                        // if it doesn't contain a wildcard then perform simple string equality
                        nonregexes.Add(pattern);
                    }
                }

                foreach (var item in items)
                {
                    string name = get_name(item);

                    bool matches = (nonregexes.Contains(name)) || (regexes.Any(regex => regex.IsMatch(name)));

                    if (action == FilterAction.Exclude && !matches)
                    {
                        yield return item;
                    }
                    else if (action == FilterAction.Include && matches)
                    {
                        yield return item;
                    }
                }
            }
        }
    }
}