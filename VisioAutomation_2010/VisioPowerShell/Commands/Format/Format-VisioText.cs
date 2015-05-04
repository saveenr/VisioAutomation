using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Format, "VisioText")]
    public class Format_VisioText : VisioCmdlet
    {

        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.ParameterAttribute(Mandatory = false)]
        [SMA.ValidateNotNullOrEmptyAttribute]
        public string  Font { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter Togglecase { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Font != null)
            {
                this.client.Text.SetFont(this.Shapes, this.Font);                
            }

            if (this.Togglecase)
            {
                this.client.Text.ToogleCase(this.Shapes);
            }
        }
    }
}