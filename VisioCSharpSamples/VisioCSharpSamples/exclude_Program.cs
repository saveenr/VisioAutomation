using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioCSharpSamples
{

    internal class Program
    {
        private static void Main(string[] args)
        {
            var app = new IVisio.Application();
            var doc = app.Documents.Add("");
            Samples.AutoConnect_ConnectorToolDataObject(doc);
        }
    }
}