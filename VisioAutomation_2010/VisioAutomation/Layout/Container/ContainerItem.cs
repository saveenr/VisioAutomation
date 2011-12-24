using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.ContainerLayout
{
    public class ContainerItem
    {
        public string Text { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public Shape VisioShape { get; set; }
        public short ShapeID { get; set; }

        public VA.ShapeSheet.FormulaLiteral FillForegnd;
        public VA.ShapeSheet.FormulaLiteral LineWeight;
        public VA.ShapeSheet.FormulaLiteral LinePattern;
        public VA.ShapeSheet.FormulaLiteral VerticalAlign;

        public ContainerItem(string text)
        {
            this.Text = text;
        }
    }
}
