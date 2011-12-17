using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "ShapeXMLDescription")]
    public class Get_ShapeXMLDescription : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.WriteObject( scriptingsession.Developer.GetXMLDescription());
        }
    }

}