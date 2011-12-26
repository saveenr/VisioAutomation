using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.Format;
using VisioAutomation.Text;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.ContainerLayout
{
    public class Container
    {
        public string Text { get; set; }
        public List<ContainerItem> ContainerItems { get; set; }
        public Shape VisioShape { get; set; }
        public VA.Drawing.Rectangle Rectangle;
        public short ShapeID;

        public Container(string text)
        {
            this.Text = text;
            this.ContainerItems = new List<ContainerItem>();
        }

        public ContainerItem Add(string text)
        {
            var ct = new ContainerItem(text);
            this.ContainerItems.Add(ct);
            return ct;
        }
    }

    public class Formatting
    {
        public VA.Format.ShapeFormatCells ShapeFormatCells;
        public VA.Text.CharacterFormatCells CharacterFormatCells;
        public VA.Text.ParagraphFormatCells ParagraphFormatCells;
        public VA.Text.TextBlockFormatCells TextBlockFormatCells;

        public Formatting()
        {
            this.ShapeFormatCells = new ShapeFormatCells();
            this.CharacterFormatCells = new CharacterFormatCells();
            this.ParagraphFormatCells = new ParagraphFormatCells();
            this.TextBlockFormatCells = new TextBlockFormatCells();
        }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid, short shapeid2)
        {
            this.CharacterFormatCells.Apply(update, shapeid, 0);
            this.ParagraphFormatCells.Apply(update, shapeid, 0);
            this.ShapeFormatCells.Apply(update, shapeid2);
            this.TextBlockFormatCells.Apply(update, shapeid);
        }
    }
}
