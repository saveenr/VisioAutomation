using VA=VisioAutomation;

namespace VisioAutomation
{
    public static class TextUtil
    {
        public static System.Text.RegularExpressions.Regex GetRegexForWildcardPattern(string cellname, bool ignorecase)
        {
            string pat = "^" + System.Text.RegularExpressions.Regex.Escape(cellname)
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
    }
}