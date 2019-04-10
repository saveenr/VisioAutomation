using System.Globalization;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioPage)]
    public class NewVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public double Width = -1.0;
        
        [SMA.Parameter(Mandatory = false)] 
        public double Height = -1.0;

        [SMA.Parameter(Mandatory = false)]
        public string Name { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public VisioPowerShell.Models.PageCells Cells { get; set; }

        protected override void ProcessRecord()
        {
            var targetdoc = new VisioScripting.TargetDocument();

            this.Client.Output.WriteVerbose("Creating a new page");
            var page = this.Client.Page.NewPage(targetdoc , null, false);
            
            if (this.Name != null)
            {
                if (this.Name.Length == 0)
                {
                    throw new System.ArgumentException("Name can't be empty");
                }

                string n = this.Name.Trim();

                if (n.Length == 0)
                {
                    throw new System.ArgumentException("Name can't be empty");
                }

                this.Client.Output.WriteVerbose("Setting page name \"{0}\"", n);
                page.NameU = n;
            }

            if (this.Width > 0 || this.Height > 0)
            {
                // width and height are used and there isn't a PageCells object
                // then create one
                this.Cells = this.Cells ?? new Models.PageCells();
                if (this.Width > 0)
                {
                    this.Cells.PageWidth = this.Width.ToString(CultureInfo.InvariantCulture);
                }
                if (this.Height > 0)
                {
                    this.Cells.PageHeight = this.Height.ToString(CultureInfo.InvariantCulture);
                }
            }

            if (this.Cells != null)
            {
                var targetpage_shapesheet = page.PageSheet;
                int targetpage_shapesheetid = targetpage_shapesheet.ID;

                var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
                writer.BlastGuards = true;
                writer.TestCircular = true;
                this.Cells.Apply(writer, (short)targetpage_shapesheetid);

                this.Client.Output.WriteVerbose("Updating Cells for new page");
                writer.Commit(page, VisioAutomation.ShapeSheet.CellValueType.Formula);
            }

            this.WriteObject(page);
        }
    }
}

