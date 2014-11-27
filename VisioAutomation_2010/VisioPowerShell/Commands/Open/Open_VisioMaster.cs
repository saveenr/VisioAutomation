namespace VisioPowerShell.Commands
{
    [System.Management.Automation.Cmdlet(System.Management.Automation.VerbsCommon.Open, "VisioMaster")]
    public class Open_VisioMaster : VisioCmdlet
    {
        [System.Management.Automation.Parameter(Position = 0, Mandatory = true)]
        public Microsoft.Office.Interop.Visio.Master Master;

        protected override void ProcessRecord()
        {
            this.client.Master.OpenForEdit(this.Master);
        }
    }
}