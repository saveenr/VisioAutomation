
namespace VisioPowerShell.Commands.VisioApplication;

[SMA.Cmdlet(SMA.VerbsCommon.Close, Nouns.VisioApplication)]
public class CloseVisioApplication : VisioCmdlet
{
    protected override void ProcessRecord()
    {
        this.Client.Application.CloseApplication();
    }
}