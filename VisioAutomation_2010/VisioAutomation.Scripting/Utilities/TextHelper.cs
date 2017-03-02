using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Utilities
{
    internal static class TextHelper
    {
        internal static string toggle_case(string input_string)
        {
            if (input_string.Length == 0)
            {
                return input_string;
            }

            string t_upper = input_string.ToUpper();
            string t_lower = input_string.ToLower();

            string output_string = null;
            if (input_string == t_upper)
            {
                output_string = t_lower;
            }
            else if (input_string == t_lower)
            {
                var cultureInfo = System.Globalization.CultureInfo.CurrentCulture;
                var textInfo = cultureInfo.TextInfo;
                var t_case = textInfo.ToTitleCase(input_string);

                output_string = t_case;
            }
            else
            {
                output_string = t_upper;
            }

            return output_string;
        }

        public static void set_text_wrapping(IVisio.Page page,
                                               IList<int> shapeids,
                                               bool wrap)
        {
            const string formula_wrap = "WIDTH*1";
            const string formula_no_wrap = "TEXTWIDTH(TheText)";
            string formula = wrap ? formula_wrap : formula_no_wrap;
            var writer = new VisioAutomation.ShapeSheet.ShapeSheetWriter();
            
            foreach (int shapeid in shapeids)
            {
                writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.TxtWidth, formula);
            }

            writer.Commit(page);
        }

        public static void Join(System.Text.StringBuilder sb, string s, IEnumerable<string> tokens)
        {
            int n = tokens.Count();
            int c = tokens.Select(t => t.Length).Sum();
            c += (n > 1) ? s.Length*n : 0;
            c += sb.Length;
            sb.EnsureCapacity(c);

            int i = 0;
            foreach (string t in tokens)
            {
                if (i > 0)
                {
                    sb.Append(s);
                }
                sb.Append(t);
                i++;
            }
        }
    }
}