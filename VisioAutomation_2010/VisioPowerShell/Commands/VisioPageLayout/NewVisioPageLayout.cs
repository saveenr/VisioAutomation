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
                var plo = new VisioAutomation.PageLayouts.FlowchartLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.Hierarchy)
            {
                var plo = new VisioAutomation.PageLayouts.HierarchyLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.Circular)
            {
                var plo = new VisioAutomation.PageLayouts.CircularLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.CompactTree)
            {
                var plo = new VisioAutomation.PageLayouts.CompactTreeLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutType == VisioPowerShell.Models.PageLayoutType.RadialLayout)
            {
                var plo = new VisioAutomation.PageLayouts.RadialLayout();
                this.WriteObject(plo);
            }
            else
            {
                throw new System.ArgumentException("Unsupported layout");
            }
        }
    }
}