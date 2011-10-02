using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace InfoGraphicsPy
{
    public class CategoryItem
    {
        public string Text;
        public List<CategoryItem> Items; 
        public CategoryItem(string s)
        {
            this.Text = s;
        }
    }
}
