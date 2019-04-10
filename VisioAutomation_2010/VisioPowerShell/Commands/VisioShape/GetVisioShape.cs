using System.Linq;
using VisioScripting;
using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName="name", Position = 0, Mandatory = false)]
        public string[] Name;

        [SMA.Parameter(ParameterSetName = "id", Position = 0, Mandatory = false)]
        public int[] Id;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Recursive;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SubSelected;

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage().Resolve(this.Client);

            // Handle the case where neither names nor ids where passed
            if (this.Name == null && this.Id == null)
            {
                // return selected shapes

                var selection = new VisioScripting.TargetSelection();

                if (this.Recursive)
                {
                    this.WriteVerbose("Returning selected shapes (nested)");
                    var shapes = this.Client.Selection.GetShapesRecursive(selection);
                    this.WriteObject(shapes, true);
                }
                if (this.SubSelected)
                {
                    this.WriteVerbose("Returning selected shapes (subselect)");
                    var shapes = this.Client.Selection.GetSubSelectedShapes(selection);
                    this.WriteObject(shapes, true);
                }
                else
                {
                    this.WriteVerbose("Returning selected shapes ");
                    var shapes = this.Client.Selection.GetSelectedShapes(selection);
                    this.WriteObject(shapes, true);
                }

                return;
            }

            // Handle the case where names where passed
            if (this.Name != null)
            {
                string str_asterisk = "*";
                if (this.Name.Contains(str_asterisk))
                {
                    var shapes = this.Client.Page.GetShapesOnPage(targetpage);
                    this.WriteObject(shapes, true);
                }
                else
                {
                    var shapes = this.Client.Page.GetShapesOnPageByName(targetpage, this.Name);
                    this.WriteObject(shapes, true);
                }

                return;
            }

            // Handle the case where ids where passed
            if (this.Id != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByID(targetpage,this.Id);
                this.WriteObject(shapes, true);

                return;
            }

            throw new System.ArgumentOutOfRangeException();
        }
    }
}