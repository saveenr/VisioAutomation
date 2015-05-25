using System;
using System.Windows.Forms;

namespace VisioPowerTools2010
{
    public partial class FormDeveloper : Form
    {
        public FormDeveloper()
        {
            this.InitializeComponent();
        }

        private void buttonHierarchy_Click(object sender, EventArgs e)
        {
            var client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
            client.Developer.DrawNamespaces();
        }

        private void buttonDiagramWithClasses_Click(object sender, EventArgs e)
        {
            var client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
            client.Developer.DrawNamespacesAndClasses();
        }

        private void buttonClassDiagrams_Click(object sender, EventArgs e)
        {
            var client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
            client.Developer.DrawScriptingDocumentation();
        }

        private void buttonEnums_Click(object sender, EventArgs e)
        {
            var client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
            client.Developer.DrawInteropEnumDocumentation();
        }
    }
}
