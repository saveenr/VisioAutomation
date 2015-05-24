using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.New
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioContainer")]
    public class New_VisioContainer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true,ParameterSetName="MasterObject")]
        public IVisio.Master Master { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "MasterName")]
        public string MasterName { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Master != null)
            {
                var shape = this.client.Master.DropContainer(this.Master);
                this.WriteObject(shape);
            }
            else if (this.MasterName != null)
            {
                var shape = this.client.Master.DropContainer(this.Master);
                this.WriteObject(shape);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
                
            }
        }
    }
}
