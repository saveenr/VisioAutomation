using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Helpers
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
                var culture_info = System.Globalization.CultureInfo.CurrentCulture;
                var text_info = culture_info.TextInfo;
                var t_case = text_info.ToTitleCase(input_string);

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
            var writer = new SidSrcWriter();
            
            foreach (int shapeid in shapeids)
            {
                writer.SetValue((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.TextXFormWidth, formula);
            }

            writer.Commit(page, VisioAutomation.ShapeSheet.CellValueType.Formula);
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