using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class NewVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public double Width = -1.0;
        
        [SMA.Parameter(Mandatory = false)] 
        public double Height = -1.0;

        [SMA.Parameter(Mandatory = false)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            var page = this.Client.Page.New(null, false);
            NewVisioPage.set_page_size(this.Client, this.Width, this.Height);
            
            if (this.Name != null)
            {
                this.Client.Page.SetName(this.Name);
            }

            this.WriteObject(page);
        }

        public static void set_page_size(VisioScripting.Client client, double width, double height)
        {
            double? w = null;
            double? h = null;

            if (width > 0)
            {
                w = width;
            }

            if (height > 0)
            {
                h = height;
            }

            if (w.HasValue || h.HasValue)
            {
                client.Page.SetSize(w, h);
            }
        }
    }
}