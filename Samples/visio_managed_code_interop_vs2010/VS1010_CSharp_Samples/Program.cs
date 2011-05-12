using IVisio = Microsoft.Office.Interop.Visio;

namespace visio_managed_code_interop_vs2010
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var app = new IVisio.ApplicationClass();
            var docs = app.Documents;
            var doc = docs.Add("");

            VS2010_CSharp_Samples.Shape_GetFormulas(doc);
            VS2010_CSharp_Samples.Shape_GetResults(doc);
            VS2010_CSharp_Samples.Page_GetFormulas(doc);
            VS2010_CSharp_Samples.Page_GetResults(doc);
            VS2010_CSharp_Samples.Shape_SetFormulas(doc);
            VS2010_CSharp_Samples.Shape_SetResults(doc);
            VS2010_CSharp_Samples.Page_SetFormulas(doc);
            VS2010_CSharp_Samples.Page_SetResults(doc);
        }
    }



}