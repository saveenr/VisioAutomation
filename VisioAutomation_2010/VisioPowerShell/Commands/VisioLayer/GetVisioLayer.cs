using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioLayer
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioLayer)]
    public class GetVisioLayer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(this.Name))
            {
                // get all
                var layers = this.Client.Layer.GetLayersOnActivePage();
                this.WriteObject(layers, true);
            }
            else
            {
                // get all that match a specific name
                var layer = this.Client.Layer.FindLayersOnActivePageByName(this.Name);
                this.WriteObject(layer);
            }
        }
    }
}