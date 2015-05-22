using System.Collections.Generic;
using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace VisioPowerTools
{
    public partial class VisioPowerToolsAddIn
    {
        private static VisioAutomation.Scripting.Client g_client;
        internal static VisioPowerTools.PowerToolsClientContext g_clientcontext;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.buttons = new List<Office.CommandBarButton>();
                vis_cmd_bars = (Office.CommandBars) this.Application.CommandBars;
                this.customize_visio_menu();
                this.addinprefs = new AddInData();
            }
            catch (System.Exception exc)
            {
                string msg = string.Format("Unhandled Exception for {0} start-up: {1}",
                                           typeof (VisioPowerToolsAddIn).Name, exc.Message);
                MessageBox.Show(msg);
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {           
        }
        
        public static VisioAutomation.Scripting.Client Client
        {
            get
            {
                if (g_client == null)
                {
                    if ( g_clientcontext == null )
                    {
                        g_clientcontext = new PowerToolsClientContext();
                    }

                    var application = Globals.VisioPowerToolsAddIn.Application;
                    g_client = new VisioAutomation.Scripting.Client(application);
                    g_client.Context = g_clientcontext;
                }
                else
                {
                    // do nothing
                }

                if (g_client.VisioApplication == null)
                {
                    throw new VisioAutomation.AutomationException("Internal Error: Unexpected null for visio application");
                }
                return g_client;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}