

namespace VisioAutomation.Models.Documents.Forms
{
    internal class FormRenderingContext
    {
        public IVisio.Application Application;
        public IVisio.Document Document;
        public IVisio.Page Page;
        public IVisio.Pages Pages;
        public Dictionary<string, int> NameToFontID;

        public IVisio.Fonts Fonts;

        public FormRenderingContext()
        {
            var compare = System.StringComparer.InvariantCultureIgnoreCase;
            this.NameToFontID = new Dictionary<string, int>(compare);
        }

        public int GetFontID(string name)
        {
            if (this.NameToFontID.ContainsKey(name))
            {
                return this.NameToFontID[name];
            }
            else
            {
                var font = this.Fonts[name];
                int id = font.ID;
                return id;
            }
        }
    }
}