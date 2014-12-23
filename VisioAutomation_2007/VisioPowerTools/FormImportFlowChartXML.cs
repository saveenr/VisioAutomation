using System.Collections.Generic;
using System.Windows.Forms;
using VA=VisioAutomation;

namespace VisioPowerTools
{
    public partial class FormImportFlowChartXML 
    {
        private static string cached_filename;
        public FormImportFlowChartXML()
        {
            InitializeComponent();
        }

        private void FormImportFlowChartXML_Load(object sender, System.EventArgs e)
        {
            if (cached_filename!=null)
            {
                this.filenamePicker1.Filename = cached_filename;
            }

        }

        private void buttonImport_Click(object sender, System.EventArgs e)
        {
            cached_filename = this.filenamePicker1.Filename;
            string filename = this.filenamePicker1.Filename;

            this.labelMessageLog.ResetText();

            if (string.IsNullOrEmpty(filename))
            {
                MessageBox.Show("Enter a filename to import");
                return;
            }

            if (!System.IO.Path.IsPathRooted(filename))
            {
                MessageBox.Show("Enter an absolute filename to import");
                return;
            }

            if (!System.IO.File.Exists(filename))
            {
                MessageBox.Show("File does not exist");
                return;
            }

            var ss = VisioPowerToolsAddIn.Client;
            System.Xml.Linq.XDocument xdoc;
            try
            {
                xdoc = System.Xml.Linq.XDocument.Load(filename);
            }
            catch (System.Xml.XmlException exc)
            {
                string msg = exc.Message + "\n" + exc.StackTrace;
                this.write_msg(msg);                
                MessageBox.Show("Failed to load XML");
                return;
            }

            IList<VA.Models.DirectedGraph.Drawing> drawings;
            try
            {
                VisioPowerToolsAddIn.g_clientcontext.OnWriteString += write_msg;
                drawings = VisioAutomation.Scripting.DirectedGraph.DirectedGraphBuilder.LoadFromXML(ss, xdoc);
                VisioPowerToolsAddIn.g_clientcontext.OnWriteString -= write_msg;
            }
            catch (VisioAutomation.AutomationException)
            {
                MessageBox.Show("Failed to Build flowchart from XML");
                return;
            }

            bool close_form = false;
            try
            {
                VisioPowerToolsAddIn.g_clientcontext.OnWriteString += write_msg;
                foreach (var d in drawings)
                {
                    var page = ss.Page.New(null, false);
                    d.Render(page);
                }
                VisioPowerToolsAddIn.g_clientcontext.OnWriteString -= write_msg;
            }
            catch(VisioAutomation.AutomationException)
            {
                MessageBox.Show("Failed to render diagram");
            }

            if (close_form)
            {
                this.Close();                
            }
        }

        private void write_msg(string s)
        {
            this.textBoxOutput.AppendText(s+"\n");
        }

        private void buttonCancel_Click(object sender, System.EventArgs e)
        {
            cached_filename = this.filenamePicker1.Filename;
            this.Close();
        }
    }
}
