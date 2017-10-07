using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioContainer)]
    public class NewVisioContainer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true,ParameterSetName="MasterObject")]
        public IVisio.Master Master { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "MasterName")]
        public string MasterName { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Master != null)
            {
                var shape = this.Client.Master.DropContainer(this.Master);
                this.WriteObject(shape);
            }
            else if (this.MasterName != null)
            {
                var shape = this.Client.Master.DropContainer(this.MasterName);
                this.WriteObject(shape);
            }
            else
            {
                string msg = string.Format("Either -{0} or -{1} must be provided.", nameof(this.Master),
                    nameof(this.MasterName));
                throw new System.ArgumentException(msg);
            }
        }
    }
}
