using System.Collections.Generic;
using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace VisioPowerTools
{
    public partial class VisioPowerToolsAddIn
    {
        private static VisioAutomation.Scripting.Client g_scripting_session;
        internal static VisioPowerTools.PowerToolsSessionOptions g_session_options;
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
        
        public static VisioAutomation.Scripting.Client ScriptingSession
        {
            get
            {
                if (g_scripting_session == null)
                {
                    if ( g_session_options == null )
                    {
                        g_session_options = new PowerToolsSessionOptions();
                    }

                    var application = Globals.VisioPowerToolsAddIn.Application;
                    g_scripting_session = new VisioAutomation.Scripting.Client(application);
                    g_scripting_session.Context = g_session_options;
                }
                else
                {
                    // do nothing
                }

                if (g_scripting_session.VisioApplication == null)
                {
                    throw new VisioAutomation.AutomationException("Internal Error: Unexpected null for visio application");
                }
                return g_scripting_session;
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