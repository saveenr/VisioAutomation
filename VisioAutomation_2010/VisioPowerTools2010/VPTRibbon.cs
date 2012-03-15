using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace VisioPowerTools2010
{
    public partial class VPTRibbon
    {
        private void VPTRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonHelp_Click_1(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Hello World");

        }

        private void buttonImportColors_Click(object sender, RibbonControlEventArgs e)
        {
            /*
            239, 62, 54
 68,198,234
 10,175,220
 13,117,144
             * 
             */
            
            var form = new FormImportColors();
            var result = form.ShowDialog();
            if (result == DialogResult.OK)
            {
                var colors = form.Colors;
                if (colors.Count < 1)
                {
                    return;
                }

                var app = Globals.ThisAddIn.Application;
                var docs = app.Documents;
                var doc = docs.Add("");
                var page = doc.Pages[1];
                double y = 8;
                double x = 1;
                foreach (var color in colors)
                {
                    var shape = page.DrawRectangle(x, y, x + 1.0, y + 1.0);
                    shape.CellsU["FillForegnd"].FormulaForceU = color.ToFormula();

                    y -= 1.5;

                }
            }

        }
    }
}
