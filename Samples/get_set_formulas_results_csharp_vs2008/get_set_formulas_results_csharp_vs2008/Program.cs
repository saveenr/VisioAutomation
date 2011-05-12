using IVisio = Microsoft.Office.Interop.Visio;

namespace DemoVisioGetSetFormulasResultsScenarios
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new IVisio.ApplicationClass();
            var docs = app.Documents;
            var doc = docs.Add("");

            CSharpSamples.Shape_GetFormulas(doc);
            CSharpSamples.Shape_GetResults(doc);
            CSharpSamples.Page_GetFormulas(doc);
            CSharpSamples.Page_GetResults(doc);
            CSharpSamples.Shape_SetFormulas(doc);
            CSharpSamples.Shape_SetResults(doc);
            CSharpSamples.Page_SetFormulas(doc);
            CSharpSamples.Page_SetResults(doc);
        }
    }
}
