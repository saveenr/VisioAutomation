using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageProperty")]
    public class Set_VisioPageProperty: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public string Width { get; set; }
        [SMA.Parameter(Mandatory = false)] public string Height { get; set; }
        [SMA.Parameter(Mandatory = false)] public string PageBottomMargin;
        [SMA.Parameter(Mandatory = false)] public string PageHeight;
        [SMA.Parameter(Mandatory = false)] public string PageLeftMargin;
        [SMA.Parameter(Mandatory = false)] public string PageLineJumpDirX;
        [SMA.Parameter(Mandatory = false)] public string PageLineJumpDirY;
        [SMA.Parameter(Mandatory = false)] public string PageRightMargin;
        [SMA.Parameter(Mandatory = false)] public string PageScale;
        [SMA.Parameter(Mandatory = false)] public string PageShapeSplit;
        [SMA.Parameter(Mandatory = false)] public string PageTopMargin;
        [SMA.Parameter(Mandatory = false)] public string PageWidth;
        [SMA.Parameter(Mandatory = false)] public IList<IVisio.Page> Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;


        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new [] { scriptingsession.Page.Get() };

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;
                var id = pagesheet.ID16;
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageHeight, this.Height);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageScale, this.PageScale);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PageWidth, this.PageWidth);
                update.Execute(page);

            }

        }
    }
}