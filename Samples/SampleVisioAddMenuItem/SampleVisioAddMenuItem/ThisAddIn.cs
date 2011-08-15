using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace SampleVisioAddMenuItem
{
    public partial class ThisAddIn
    {

        Office.CommandBarButton new_button;
        Office.CommandBars vis_cmd_bars;
        Office.CommandBarPopup vis_file_menu;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            vis_cmd_bars = this.Application.CommandBars as Office.CommandBars;

            vis_file_menu = vis_cmd_bars["Menu Bar"].Controls["&File"] as Office.CommandBarPopup;

            new_button = (Office.CommandBarButton) vis_file_menu.Controls.Add(
                Office.MsoControlType.msoControlButton, // Type
                this.missing, // Object
                this.missing, // Id
                2, // Before
                true // Temporary
                ) ;

            new_button.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
            new_button.Caption = "My New Menu Item";
            //new_button.Tag = "My New Menu Item";
            new_button.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(btnOpen_Click);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void btnOpen_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {

            System.Windows.Forms.MessageBox.Show("You pressed me!");

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
