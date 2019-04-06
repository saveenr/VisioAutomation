using SMA = System.Management.Automation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioLayer
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioLayer)]
    public class GetVisioLayer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name;

        [SMA.Parameter(Position = 0, Mandatory = false)]
        public IVisio.Page Page;

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage();
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(this.Name))
            {
                // get all
                var layers = this.Client.Layer.GetLayersOnPage(targetpage);
                this.WriteObject(layers, true);
            }
            else
            {
                // get all that match a specific name
                var layer = this.Client.Layer.FindLayersOnPageByName(targetpage, this.Name);
                this.WriteObject(layer);
            }
        }
    }
}