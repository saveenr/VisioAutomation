

namespace VisioScripting;

public class TargetPage : TargetObject<IVisio.Page>
{

    private TargetPage() : base()
    {
    }

    public TargetPage(IVisio.Page page) : base(page)
    {
    }

    public TargetPage ResolveToPage(Client client)
    {
        if (this.Resolved)
        {
            return this;
        }

        var cmdtarget = client.GetCommandTarget(CommandTargetFlags.RequirePage);

        client.Output.WriteVerbose("Resolving to active page (name={0})", cmdtarget.ActivePage.Name);

        return new TargetPage(cmdtarget.ActivePage);
    }

    public IVisio.Page Page => this._get_item_safe();

    public static TargetPage Auto => new TargetPage();
}