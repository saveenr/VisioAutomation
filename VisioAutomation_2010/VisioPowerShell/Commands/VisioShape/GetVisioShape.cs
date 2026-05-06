using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    // Parameter sets:
    //   "active"       -> -ActiveSelection switch: return the current selection.
    //   "shapebyid"    -> -ID <int[]>:           return shapes on the page with those IDs.
    //   "shapebyname"  -> -Name <string[]>:      return shapes on the page with those names.
    //                                            Also the DEFAULT set: a no-args call lands
    //                                            here with Name == null and returns every
    //                                            shape on the (resolved) page.
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape, DefaultParameterSetName = "shapebyname")]
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
                this.WriteVerbose("Returning selected shapes");
                var shapes_selected = this.Client.Selection.GetSelectedShapes(VisioScripting.TargetSelection.Auto);
                this.WriteObject(shapes_selected, true);
                return;
            }

            var targetpage = new VisioScripting.TargetPage(this.Page);

            if (this.ID != null)
            {
                var shapes_by_id = this.Client.Page.GetShapesOnPageByID(targetpage, this.ID);
                this.WriteObject(shapes_by_id, true);
                return;
            }

            if (this.Name != null)
            {
                var shapes_by_name = this.Client.Page.GetShapesOnPageByName(targetpage, this.Name);
                this.WriteObject(shapes_by_name, true);
                return;
            }

            // No filter (default parameter set, with -Name unset): return every shape on the resolved page.
            var shapes_on_page = this.Client.Page.GetShapesOnPage(targetpage);
            this.WriteObject(shapes_on_page, true);
        }
    }
}