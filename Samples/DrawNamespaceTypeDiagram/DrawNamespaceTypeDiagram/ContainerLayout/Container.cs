using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomationMin;

namespace ContainerLayout
{
    class Container
    {
        public string Text { get; set; }
        public List<ContainerItem> ContainerItems { get; set; }
        public Shape VisioShape { get; set; }
        public VAM.Rectangle Rectangle;
        public short ShapeID;

        public VAM.FormulaLiteral FillForegnd;
        public VAM.FormulaLiteral LineWeight;
        public VAM.FormulaLiteral LinePattern;
        public VAM.FormulaLiteral VerticalAlign;

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
}
