using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void HelloWorld()
        {
            var app = new IVisio.ApplicationClass();
            var docs = app.Documents;
            var doc = docs.Add("");

            var page = app.ActivePage;
            var shape0 = page.DrawRectangle(1, 2, 6, 3);
            shape0.Text = "Hello World";
        }
    }
}