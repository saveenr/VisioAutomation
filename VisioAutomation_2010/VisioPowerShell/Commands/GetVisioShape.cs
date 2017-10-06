using System.Linq;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName="name", Position = 0, Mandatory = false)]
        public string[] Name;

        [SMA.Parameter(ParameterSetName = "id", Position = 0, Mandatory = false)]
        public int[] Id;
        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Recursive;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SubSelected;

        protected override void ProcessRecord()
        {
            // Handle the case where neither names nor ids where passed
            if (this.Name == null && this.Id == null)
            {
                // return selected shapes

                if (this.Recursive)
                {
                    this.WriteVerbose("Returning selected shapes (nested)");
                    var shapes = this.Client.Selection.GetShapesRecursive();
                    this.WriteObject(shapes, false);
                }
                if (this.SubSelected)
                {
                    this.WriteVerbose("Returning selected shapes (subselecte)");
                    var shapes = this.Client.Selection.GetSubSelectedShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    this.WriteVerbose("Returning selected shapes ");
                    var shapes = this.Client.Selection.GetShapes();
                    this.WriteObject(shapes, false);
                }

                return;
            }

            // Handle the case where names where passed
            if (this.Name != null)
            {
                string str_asterisk = "*";
                if (this.Name.Contains(str_asterisk))
                {
                    var shapes = this.Client.Draw.GetAllShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    var strings = this.Name.Where(i => i is string).Cast<string>().ToArray();
                    var shapes = this.Client.Page.GetShapesByName(strings);
                    this.WriteObject(shapes, false);
                }

                return;
            }

            // Handle the case where ids where passed
            if (this.Id != null)
            {
                var shapes = this.Client.Page.GetShapesByID(this.Id);
                this.WriteObject(shapes, false);

                return;
            }

            throw new System.ArgumentOutOfRangeException();
        }
    }
}