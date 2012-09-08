using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio=Microsoft.Office.Interop.Visio;
using SXL = System.Xml.Linq;

namespace VisioAutomation
{
    public static class ExportHelper
    {
        public static void ExportSelectionAsSVGXHTML2(IVisio.Selection selection, string filename, System.Action<string> verboselog)
        {
            // Save temp SVG
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";
            selection.Export(svg_filename);

            // Load temp SVG
            var load_svg_timer = new System.Diagnostics.Stopwatch();
            var svg_doc = SXL.XDocument.Load(svg_filename);
            load_svg_timer.Stop();
            verboselog(string.Format("Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds));

            // Delete temp SVG
            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }
            else
            {
                // TODO: throw an exception
            }

            verboselog(string.Format("Creating XHTML with embedded SVG"));
            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                verboselog(string.Format("Deleting \"{0}\"", filename));
                System.IO.File.Delete(filename);
            }

            var xhtml_doc = new SXL.XDocument();
            var xhtml_root = new SXL.XElement("{http://www.w3.org/1999/xhtml}html");
            xhtml_doc.Add(xhtml_root);
            var svg_node = svg_doc.Root;
            svg_node.Remove();

            var body = new SXL.XElement("{http://www.w3.org/1999/xhtml}body");
            xhtml_root.Add(body);
            body.Add(svg_node);

            xhtml_doc.Save(filename);
            verboselog(string.Format("Done writing XHTML file \"{0}\"", filename));
        }


        public static void ExportSelectionAsXAML2(IVisio.Selection sel, string filename, System.Action<string> verboselog)
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
            var convert_timer = new System.Diagnostics.Stopwatch();

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
