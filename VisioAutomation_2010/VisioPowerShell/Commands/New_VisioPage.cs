using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class New_VisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false)] 
        public double Width = -1.0;
        
        [Parameter(Mandatory = false)] 
        public double Height = -1.0;

        [Parameter(Mandatory = false)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            var page = this.Client.Page.New(null, false);
            New_VisioPage.set_page_size(this.Client, this.Width, this.Height);
            
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