using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VisioPowerTools2010
{
    public partial class FormDirectedGraph : Form
    {
        public FormDirectedGraph()
        {
            InitializeComponent();
        }

        public string GraphText
        {
            get { return this.textBox1.Text; }
        }
    }
}
