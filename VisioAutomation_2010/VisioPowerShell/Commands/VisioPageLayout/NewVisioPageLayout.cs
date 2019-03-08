using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{

    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioPageLayout)]
    public class NewVisioPageLayout : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public VisioPowerShell.Models.PageLayoutType LayoutType { get; set; }


        protected override void ProcessRecord()
        {
            if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.FlowChart)
            {
                var plo = new VisioAutomation.LayoutStyles.FlowchartLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.Hierarchy)
            {
                var plo = new VisioAutomation.LayoutStyles.HierarchyLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.Circular)
            {
                var plo = new VisioAutomation.LayoutStyles.CircularLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.CompactTree)
            {
                var plo = new VisioAutomation.LayoutStyles.CompactTreeLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.RadialLayout)
            {
                var plo = new VisioAutomation.LayoutStyles.RadialLayoutStyle();
                this.WriteObject(plo);
            }
            else
            {
                throw new System.ArgumentException("Unsupported layout");
            }
        }
    }
}