using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void Shape_FormatText(IVisio.Document doc)
        {
            var page = Util.CreateStandardPage(doc, "SSR");
            var shape = Util.CreateStandardShape(page);

            shape.Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            // create a new character row
            var chars = shape.Characters;
            chars.Begin = 15;
            chars.End = 25;

            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;

           chars.CharProps[(short)IVisio.VisCellIndices.visCharacterColor] = (short)0;
            var rownum = chars.CharPropsRow[(short) default_chars_bias];

            shape.Cells["Char.Color[" + (rownum+1) + "]"].Formula = "rgb(255,0,0)";

        }
    }
}