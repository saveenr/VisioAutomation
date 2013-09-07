using System.Collections.Generic;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class FormRenderingContext
    {
        public IVisio.Application Application;
        public IVisio.Document Document;
        public IVisio.Page Page;
        public IVisio.Pages Pages;
        public Dictionary<string, int> NameToFontID;

        public IVisio.Fonts Fonts;

        public FormRenderingContext(IVisio.Application app)
        {
            this.NameToFontID = new Dictionary<string, int>(System.StringComparer.InvariantCultureIgnoreCase);
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