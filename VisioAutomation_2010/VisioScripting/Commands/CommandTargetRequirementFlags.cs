namespace VisioScripting.Commands
{
    [System.Flags]
    public enum CommandTargetRequirementFlags
    {
        RequireApplication,
        RequireActiveDocument,
        RequirePage
    }
}