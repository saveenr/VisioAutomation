using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.DOM;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    public class CategoryCell
    {
        public CategoryItem Item;
        public string XCategory;
        public string YCategory;
 
        public CategoryCell(string text, string x, string y)
        {
            this.Item = new CategoryItem(text);
            this.XCategory = x;
            this.YCategory = y;
        }
    }
}
