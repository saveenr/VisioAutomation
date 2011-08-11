using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class RenderContext
    {
        private IVisio.Fonts Fonts;
        public IVisio.Page Page;
        public VA.Drawing.Point CurrentUpperLeft;
        public double PageWidth;
        public string DefaultFont = "Segoe UI";
        public VA.Drawing.ColorRGB TileColor = new ColorRGB(0xe0e0e0);
        public VA.Drawing.ColorRGB LineColor = new ColorRGB(0xc0c0c0);

        private IDictionary<string, short> map_name_to_fontid;
 
        public VA.Drawing.ColorRGB TileColorReal = new ColorRGB(0xf0f0f0);

        public RenderContext()
        {
            this.map_name_to_fontid = new Dictionary<string, short>();
        }

        public VA.Format.ShapeFormatCells GetDefaultBkfmt()
        {
            var bkfmt = new VA.Format.ShapeFormatCells();
            bkfmt.FillForegnd = this.TileColorReal.ToFormula();
            bkfmt.LinePattern = 0;
            bkfmt.LineWeight = 0; //  VA.Convert.PointsToInches(1.0);
            bkfmt.LineColor = 0; //this.LineColor.ToFormula();
            return bkfmt;
        }

        public int GetFontID(string name)
        {
            if (this.map_name_to_fontid.ContainsKey(name))
            {
                return this.map_name_to_fontid[name];
            }
            else
            {
                short id = this.GetFont(name).ID16;
                this.map_name_to_fontid[name] = id;
                return id;
            }
        }

        public IVisio.Font GetFont(string name)
        {
            var doc = this.Page.Document;
            var fonts = doc.Fonts;
            var myfont = VA.Text.TextHelper.TryGetFont(fonts, name);
            if (myfont == null)
            {
                myfont = fonts["Arial"];
            }
            return myfont;
        }
    }
}
