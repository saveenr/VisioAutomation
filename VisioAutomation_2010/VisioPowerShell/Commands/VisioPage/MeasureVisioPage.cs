

namespace VisioPowerShell.Commands.VisioPage;

[SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioPage)]
public class MeasureVisioPage : VisioCmdlet
{
    /*
    [SMA.Parameter(Mandatory = false)]
    public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }
    */

    // CONTEXT:PAGES
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Page[] Page;

    protected override void ProcessRecord()
    {

        var targetpages = new VisioScripting.TargetPages(this.Page).ResolveToPages(this.Client);

        if (targetpages.Pages.Count < 1)
        {
            return;
        }

        var list_pagedim = VisioScripting.Models.PageDimensions.Get_PageDimensions(targetpages.Pages);

        this.WriteObject(list_pagedim,true);

    }
}