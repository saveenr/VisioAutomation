using IVisio = Microsoft.Office.Interop.Visio;
using MOC = Microsoft.Office.Core;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;

namespace VisioPowerTools2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion


        private static VAS.Session g_scripting_session;

        public static VAS.Session ScriptingSession
        {
            get
            {
                if (g_scripting_session == null)
                {
                    var application = Globals.ThisAddIn.Application;
                    g_scripting_session = new VAS.Session(application);
                }
                else
                {
                    // do nothing
                }

                if (g_scripting_session.Application == null)
                {
                    throw new VA.AutomationException("Internal Error: Unexpected null for visio application");
                }
                return g_scripting_session;
            }
        }
    }
}
