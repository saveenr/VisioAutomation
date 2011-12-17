using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "SRC")]
    public class Get_SRC : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Name;
        
        protected override void ProcessRecord()
        {
            foreach (var name in Name)
            {
                var v = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName(name);
                if (v.HasValue)
                {
                    this.WriteObject(v.Value);
                }
                else
                {
                    string msg = string.Format("Finding Cell SRC for {0} is not supported", name);
                    throw new System.ArgumentException(msg);
                }
            }
        }
    }
}