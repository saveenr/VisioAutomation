using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPage")]
    public class Set_VisioPage : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Name")]
        public string Name { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Page")]
        public IVisio.Page Page  { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Flags")]
        public VA.Scripting.PageNavigation Flag { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Name != null)
            {

                var app = this.ScriptingSession.VisioApplication;
                var doc = app.ActiveDocument;
                this.WriteVerboseEx("Retrieving Page \"{0}\"", this.Name);
                var pages = doc.Pages;
                var page = pages[this.Name];
                this.WriteVerboseEx("Setting Active Page to \"{0}\"", this.Name);
                var window = app.ActiveWindow;
                window.Page = page;
            }
            else if (this.Page != null)
            {

                var app = this.ScriptingSession.VisioApplication;
                this.WriteVerboseEx("Setting Active Page to \"{0}\"", Page.Name);
                var window = app.ActiveWindow;
                window.Page = this.Page;
            }
            else
            {
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Page.GoTo(this.Flag);
                
            }
        }
    }
}