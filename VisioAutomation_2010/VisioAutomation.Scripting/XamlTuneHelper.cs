using System.Diagnostics;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.XamlTune
{
    public static class XamlTuneHelper
    {
        public static void ExportSelectionAsXAML(Session scripting_session, string filename)
        {
            ExportSelectionAsXAML2(scripting_session.Selection.Get(), filename, s=>scripting_session.Output.Write(OutputStream.Verbose,s));
        }

        public static void ExportSelectionAsXAML2(IVisio.Selection sel, string filename, System.Action<string> verboselog)
        {
            // Save temp SVG
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";
            sel.Export(svg_filename);

            // Load temp SVG
            var load_svg_timer = new Stopwatch();
            string input_svg_content = System.IO.File.ReadAllText(svg_filename);
            load_svg_timer.Stop();
            verboselog(string.Format("Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds));

            // Delete temp SVG
            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }
            else
            {
                //TODO: Throw an Exception
            }

            verboselog(string.Format("Creating XHTML with embedded SVG"));
            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                verboselog(string.Format("Deleting \"{0}\"", filename));
                System.IO.File.Delete(filename);
            }

            verboselog(string.Format("Converting to XAML ..."));
            var convert_timer = new Stopwatch();

            string xaml;
            try
            {
                xaml = XamlTuneConverter.Svg2Xaml.ConvertFromSVG(input_svg_content);
            }
            catch (System.Exception e)
            {
                string msg = System.String.Format("Failed to convert to XAML \"{0}\"", e.Message + e.StackTrace);
                verboselog(msg);
                return;
            }
            convert_timer.Stop();

            verboselog(string.Format("Writing XAML File"));
            System.IO.File.WriteAllText(filename, xaml);
            verboselog(string.Format("Finished writing XAML File"));

            verboselog(string.Format("Finished XAML export ({0} seconds)", convert_timer.Elapsed.TotalSeconds));
        }

    }
}