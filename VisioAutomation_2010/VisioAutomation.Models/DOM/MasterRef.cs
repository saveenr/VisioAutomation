namespace VisioAutomation.Models.Dom;

public class MasterRef
{
    public string MasterName { get; private set; }
    public string StencilName { get; }
    public IVisio.Master VisioMaster { get; internal set; }

    public MasterRef(IVisio.Master master)
    {
        this.VisioMaster = master ?? throw new System.ArgumentNullException(nameof(master));
        this.MasterName = null;
        this.StencilName = null;
    }

    public MasterRef(string mastername, string stencilname)
    {
        if (mastername == null)
        {
            throw new System.ArgumentNullException(nameof(mastername));
        }

        if (MasterRef._has_stencil_extension(mastername))
        {
            throw new System.ArgumentException("Master name ends with .VSS or .VSSX");

            // Passing in the stencil name for the master name is a very common error.
            // so we make sure to check for it
        }

        if (this.StencilName != null && (!MasterRef._has_stencil_extension(stencilname)))
        {                    
            throw new System.ArgumentException("Stencil name does not end with .VSS");

            // Passing in the master name for the stencil name is a very common error.
            // so we make sure to check for it
        }

        // NOTE: Stencil names are allowed to be null. In this case 
        // it means look for the stencil in the active document

        this.VisioMaster = null;
        this.MasterName = mastername;
        this.StencilName = stencilname;
    }

    private static bool _has_stencil_extension(string s)
    {
        return _ends_with_case_insensitive(s,".vss") || _ends_with_case_insensitive(s,".vssx");
    }

    private static bool _ends_with_case_insensitive(string s, string pat)
    {
        return s.EndsWith(pat, System.StringComparison.InvariantCultureIgnoreCase);
    }
}