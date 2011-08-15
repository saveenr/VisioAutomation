using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomationMin;

namespace ContainerLayout
{
    class ContainerItem
    {
        public string Text { get; set; }
        public VAM.Rectangle Rectangle { get; set; }
        public Shape VisioShape { get; set; }
        public short ShapeID { get; set; }

        public VAM.FormulaLiteral FillForegnd;
        public VAM.FormulaLiteral LineWeight;
        public VAM.FormulaLiteral LinePattern;
        public VAM.FormulaLiteral VerticalAlign;

        public ContainerItem(string text)
        {
            this.Text = text;
        }
    }
}
