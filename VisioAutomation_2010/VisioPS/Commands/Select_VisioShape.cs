using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Select", "VisioShape")]
    public class Select_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public IVisio.Shape[] Shape;
        [SMA.Parameter(Mandatory = false)] public int[] ID;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if ( Shape !=null)
            {
                scriptingsession.Selection.Select(Shape);
            }

            if (ID!=null)
            {
                scriptingsession.Selection.Select(ID);
            }
        }
    }
}