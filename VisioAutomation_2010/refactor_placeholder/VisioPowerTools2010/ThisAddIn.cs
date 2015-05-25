using Office = Microsoft.Office.Core;

namespace VisioPowerTools2010
{
    public partial class ThisAddIn
    {
        public VisioAutomation.Scripting.Client Client;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += this.ThisAddIn_Startup;
            this.Shutdown += this.ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
