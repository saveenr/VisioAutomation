using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{

    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.LayoutStyle)]
    public class NewLayoutStyle : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public VisioPowerShell.Models.LayoutStyleType LayoutStyle { get; set; }

        protected override void ProcessRecord()
        {
            if (this.LayoutStyle == VisioPowerShell.Models.LayoutStyleType.FlowChart)
            {
                var plo = new VisioAutomation.Models.LayoutStyles.FlowchartLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutStyle == VisioPowerShell.Models.LayoutStyleType.Hierarchy)
            {
                var plo = new VisioAutomation.Models.LayoutStyles.HierarchyLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutStyle == VisioPowerShell.Models.LayoutStyleType.Circular)
            {
                var plo = new VisioAutomation.Models.LayoutStyles.CircularLayoutStyle();
                this.WriteObject(plo);
            }
            else if (this.LayoutStyle == VisioPowerShell.Models.LayoutStyleType.CompactTree)
            {
                var plo = new VisioAutomation.Models.LayoutStyles.CompactTreeLayout();
                this.WriteObject(plo);
            }
            else if (this.LayoutStyle == VisioPowerShell.Models.LayoutStyleType.RadialLayout)
            {
                var plo = new VisioAutomation.Models.LayoutStyles.RadialLayoutStyle();
                this.WriteObject(plo);
            }
            else
            {
                throw new System.ArgumentException("Unsupported layout");
            }
        }
    }
}