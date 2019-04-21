using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveSelection;

        [SMA.Parameter(Mandatory = false, Position = 0)]
        public string[] Name;

        [SMA.Parameter(Mandatory = false)]
        public int[] ID;

        // CONTEXT:PAGE
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        protected override void ProcessRecord()
        {
            if (this.ActiveSelection)
            {
                // If we arrive here then it just means get the selected shapes
                this.WriteVerbose("Returning selected shapes ");
                var selected_shapes = this.Client.Selection.GetSelectedShapes(VisioScripting.TargetSelection.Auto);
                this.WriteObject(selected_shapes, true);
                return;
            }

            // If the active selection is not specified then work on all the shapes in a page (user-specified or auto)

            var targetpage = new VisioScripting.TargetPage(this.Page);

            // First, the ID case
            if (this.ID != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByID(targetpage, this.ID);
                this.WriteObject(shapes, true);
                return;
            }



            // Handle the case where names were passed
            if (this.Name != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByName(targetpage, this.Name);
                this.WriteObject(shapes, true);

                return;
            }

            // Handle the case where ids were passed
            if (this.ID != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByID(targetpage, this.ID);
                this.WriteObject(shapes, true);
                return;
            }

            if (this.Page != null)
            {
                var shapes = this.Client.Page.GetShapesOnPage(targetpage);
                this.WriteObject(shapes, true);
                return;
            }

        }
    }
}