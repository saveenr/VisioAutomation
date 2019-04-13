using System.Linq;
using VisioScripting;
using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string[] Name;

        [SMA.Parameter(Mandatory = false)]
        public int[] Id;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Recursive;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SubSelected;

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

            // Handle the case where ids were passed
            if (this.Id != null)
            {
                var shapes = this.Client.Page.GetShapesOnPageByID(targetpage, this.Id);
                this.WriteObject(shapes, true);

                return;
            }

            var selection = new VisioScripting.TargetSelection();

            if (this.SubSelected || this.Recursive)
            {
                if (this.Name != null && this.Id != null && this.Page != null)
                {
                    throw new System.ArgumentException("SubSelect and Recursive cannot be used when the Name, Id, or Page is used");
                }

                if (this.Recursive && this.SubSelected)
                {
                    throw new System.ArgumentException("SubSelect and Recursive cannot be used together");
                }


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

            }

            // If we arrive here then it just means get the selected shapes
            this.WriteVerbose("Returning selected shapes ");
            var selected_shapes = this.Client.Selection.GetSelectedShapes(selection);
            this.WriteObject(selected_shapes, true);
        }
    }
}