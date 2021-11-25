
namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, ParameterSetName = "active")]
        public SMA.SwitchParameter ActiveSelection;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "shapebyname")]
        public string[] Name;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "shapebyid")]
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
                var shapes_selected = this.Client.Selection.GetSelectedShapes(VisioScripting.TargetSelection.Auto);
                this.WriteObject(shapes_selected, true);
                return;
            }

            // If the active selection is not specified then work on all the shapes in a page (user-specified or auto)

            var targetpage = new VisioScripting.TargetPage(this.Page);

            // First, the ID case
            if (this.ID != null)
            {
                var shapes_by_id = this.Client.Page.GetShapesOnPageByID(targetpage, this.ID);
                this.WriteObject(shapes_by_id, true);
                return;
            }

            // Then, handle the name case
            if (this.Name != null)
            {
                var shapes_by_name = this.Client.Page.GetShapesOnPageByName(targetpage, this.Name);
                this.WriteObject(shapes_by_name, true);
                return;
            }

            // Finally return all the shapes on the page

            var shapes_on_page = this.Client.Page.GetShapesOnPage(targetpage);
            this.WriteObject(shapes_on_page, true);
            return;

        }
    }
}