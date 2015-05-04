using VisioAutomation.Shapes.UserDefinedCells;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioUserDefinedCell")]
    public class Set_VisioUserDefinedCell : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)] 
        public string Prompt;

        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes; 

        protected override void ProcessRecord()
        {
            var userprop = new UserDefinedCell(this.Name, this.Value);
            if (this.Prompt != null)
            {
                userprop.Prompt = this.Prompt;
            }

            this.client.UserDefinedCell.Set(this.Shapes, userprop);
        }
    }
}