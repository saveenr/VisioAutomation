using System.Collections.Generic;
using System.Linq;
using System.Threading;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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

        internal static void set_text_wrapping(IVisio.Page page,
                                               IList<int> shapeids,
                                               bool wrap)
        {
            const string formula_wrap = "WIDTH*1";
            const string formula_no_wrap = "TEXTWIDTH(TheText)";
            string formula = wrap ? formula_wrap : formula_no_wrap;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.TxtWidth, formula);
            }

            update.Execute(page);
        }

        /// <summary>
        /// Clears the text formatting for a set of shapes
        /// </summary>
        /// <param name="shapes"></param>
        /// <param name="fontid"></param>
        /// <param name="size"></param>
        /// <param name="color"></param>
        internal static void reset_character_formatting(
            IList<IVisio.Shape> shapes, 
            int fontid, 
            double size,
            int color)
        {
            if (shapes.Count < 1)
            {
                return;
            }

            foreach (var shape in shapes)
            {
                string t = shape.Text;
                shape.Text = t;
            }

            var first_shape = shapes[0];
            var application = first_shape.Application;
            var page = application.ActivePage;
            var shapeids = shapes.Select(s => s.ID).ToList();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            
            foreach (var shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Char_Font, fontid);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Char_Size, size);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Char_Color, color);
            }

            update.Execute(page);
        }
    }
}