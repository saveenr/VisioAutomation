using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VisioAutomationDevTool
{
    public partial class FormTextWindow : Form
    {
        public FormTextWindow()
        {
            InitializeComponent();
        }

        public void SetText(IList<string> lines)
        {
                
            foreach (var line in lines)
            {
                this.textBox1.Lines = lines.ToArray();
            }
        }

    }
}
