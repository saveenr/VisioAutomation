using System.Collections.Generic;
using System.Linq;
using System.Threading;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    internal static class TextCommandsUtil
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
                var cultureInfo = Thread.CurrentThread.CurrentCulture;
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
            var update = new ShapeSheet.Update();
            
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.TxtWidth, formula);
            }

            update.Execute(page);
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