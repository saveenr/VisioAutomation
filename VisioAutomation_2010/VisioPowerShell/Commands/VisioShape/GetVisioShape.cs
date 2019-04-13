using System.Linq;
using VisioScripting;
using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public string[] Name;

        [SMA.Parameter(Mandatory = false)]
        public int[] Id;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage();

            // Name and Id cannot be used together
            if (this.Name != null && this.Id != null)
            {
                throw new System.ArgumentException("Name and ID cannot be used together");
            }

            // Handle the case where names were passed
            if (this.Name != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByName(targetpage, this.Name);
                this.WriteObject(shapes, true);

                return;
            }

            // Handle the case where ids were passed
            if (this.Id != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByID(targetpage, this.Id);
                this.WriteObject(shapes, true);

                return;
            }

            var selection = new VisioScripting.TargetSelection();

            // If we arrive here then it just means get the selected shapes
            this.WriteVerbose("Returning selected shapes ");
            var selected_shapes = this.Client.Selection.GetSelectedShapes(selection);
            this.WriteObject(selected_shapes, true);
        }
    }
}