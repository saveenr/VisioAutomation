namespace VisioExportPagesToDocs
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                System.Console.WriteLine("Syntax is: VisioExportPagesToDocs <filename.vsd> [<outptufolder>]");
                System.Environment.Exit(0);
            }

            string input_filename = args[0];
            input_filename = System.IO.Path.GetFullPath(input_filename);

            var visioapp = new Microsoft.Office.Interop.Visio.Application();
            var docs = visioapp.Documents;
            Microsoft.Office.Interop.Visio.Document doc = null;
            try
            {
                doc = docs.Open(input_filename);

                var settings = new ExporterSettings();
                settings.InputDocument = doc;
                if (args.Length >= 2)
                {
                    settings.DestinationPath = args[1];
                }
                else
                {
                    settings.DestinationPath = System.IO.Path.GetDirectoryName(input_filename);
                }

                var exporter = new Exporter(settings);
                exporter.Run();

            }
            catch (System.Runtime.InteropServices.COMException comexc)
            {
                throw new System.ArgumentException(string.Format("Failed to open file: {0}", comexc.Message));
            }
        }
    }
}
