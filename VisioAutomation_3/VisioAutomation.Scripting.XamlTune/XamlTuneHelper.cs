using System.Diagnostics;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.XamlTune
{
    public static class XamlTuneHelper
    {
        public static void ExportSelectionAsXAML(Session scripting_session, string filename)
        {
            // Save temp SVG
            string svg_filename = scripting_session.Export.SaveSelectionAsTemporarySVG();

            // Load temp SVG
            var load_svg_timer = new Stopwatch();
            string input_svg_content = System.IO.File.ReadAllText(svg_filename);
            load_svg_timer.Stop();
            scripting_session.Output.Write(OutputStream.Verbose, "Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds);

            // Delete temp SVG
            scripting_session.Export.DeleteTemporarySVG(svg_filename);

            scripting_session.Output.Write(OutputStream.Verbose, "Creating XHTML with embedded SVG");
            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                scripting_session.Output.Write(OutputStream.Verbose, "Deleting \"{0}\"", filename);
                System.IO.File.Delete(filename);
            }

            scripting_session.Output.Write(OutputStream.Verbose, "Converting to XAML ...");
            var convert_timer = new Stopwatch();

            string xaml;
            try
            {
                xaml = XamlTuneConverter.Svg2Xaml.ConvertFromSVG(input_svg_content);
            }
            catch (System.Exception e)
            {
                string msg = System.String.Format("Failed to convert to XAML \"{0}\"", e.Message + e.StackTrace);
                scripting_session.Output.Write(OutputStream.Error, msg);
                return;
            }
            convert_timer.Stop();

            scripting_session.Output.Write(OutputStream.Verbose, "Writing XAML File");
            System.IO.File.WriteAllText(filename, xaml);
            scripting_session.Output.Write(OutputStream.Verbose, "Finished writing XAML File");

            scripting_session.Output.Write(OutputStream.Verbose,
                  System.String.Format("Finished XAML export ({0} seconds)", convert_timer.Elapsed.TotalSeconds));
        }

    }
}