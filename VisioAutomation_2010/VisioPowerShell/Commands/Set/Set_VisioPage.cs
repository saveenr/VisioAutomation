using VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioPage")]
    public class Set_VisioPage : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true, ParameterSetName = "Name")]
        public string Name { get; set; }

        [SMA.ParameterAttribute(Position = 0, Mandatory = true, ParameterSetName = "Page")]
        public IVisio.Page Page  { get; set; }

        [SMA.ParameterAttribute(Position = 0, Mandatory = true, ParameterSetName = "PageNumber")]
        public int PageNumber = -1;

        [SMA.ParameterAttribute(Position = 0, Mandatory = true, ParameterSetName = "Flags")]
        public PageDirection Direction { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Name != null)
            {
                this.client.Page.Set(this.Name);
            }
            else if (this.Page != null)
            {
                this.client.Page.Set(this.Page);
            }
            else if (this.PageNumber > 0)
            {
                this.client.Page.Set(this.PageNumber);
            }
            else
            {
                this.client.Page.GoTo(this.Direction);                
            }
        }
    }
}