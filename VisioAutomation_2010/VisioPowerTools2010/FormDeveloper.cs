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
    public partial class FormDeveloper : Form
    {
        public FormDeveloper()
        {
            InitializeComponent();
        }

        private void buttonHierarchy_Click(object sender, EventArgs e)
        {
            var session = new VisioAutomation.Scripting.Session(Globals.ThisAddIn.Application);
            session.Developer.DrawVANamespaces();
        }

        private void buttonDiagramWithClasses_Click(object sender, EventArgs e)
        {
            var session = new VisioAutomation.Scripting.Session(Globals.ThisAddIn.Application);
            session.Developer.DrawVANamespacesAndClasses();

        }
    }
}
