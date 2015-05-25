using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools2010
{
    public static class ExportExtensions
    {

        public static void ExportSelectionToXAML(VA.Scripting.Client client, string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!client.Selection.HasShapes())
            {
                return;
            }

            var selection = client.Selection.Get();
            ExportExtensions.ExportSelectionToXAML(selection, filename, s => client.Output.WriteVerbose(s));
        }

        public static void ExportSelectionToXAML(IVisio.Selection sel, string filename, System.Action<string> verboselog)
        {
            // Save temp SVG
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";
            sel.Export(svg_filename);

            // Load temp SVG
            var load_svg_timer = new System.Diagnostics.Stopwatch();
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
                string msg = string.Format("Temporary SVG file could not be found: \"{0}\"", svg_filename);
                throw new VisioAutomation.Scripting.VisioOperationException(msg);
            }

            verboselog("Creating XHTML with embedded SVG");

            if (System.IO.File.Exists(filename))
            {
                verboselog(string.Format("Deleting \"{0}\"", filename));
                System.IO.File.Delete(filename);
            }

            verboselog("Converting to XAML ...");
            var convert_timer = new System.Diagnostics.Stopwatch();

            string xaml;
            try
            {
                xaml = XamlTuneConverter.Svg2Xaml.ConvertFromSVG(input_svg_content);
            }
            catch (System.Exception e)
            {
                string msg = string.Format("Failed to convert to XAML \"{0}\"", e.Message + e.StackTrace);
                verboselog(msg);
                return;
            }
            convert_timer.Stop();

            verboselog("Writing XAML File");
            System.IO.File.WriteAllText(filename, xaml);
            verboselog("Finished writing XAML File");

            verboselog(string.Format("Finished XAML export ({0} seconds)", convert_timer.Elapsed.TotalSeconds));
        }

    }

}