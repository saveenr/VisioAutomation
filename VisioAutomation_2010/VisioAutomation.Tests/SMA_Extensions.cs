namespace VisioAutomation_Tests.PowerShell
{
    public static class SMA_Extensions
    {
        public static void AddParameter(this System.Management.Automation.Runspaces.Command cmd, string name, object value)
        {
            var parameter= new System.Management.Automation.Runspaces.CommandParameter(name, value);
            cmd.Parameters.Add(parameter);            
        }    
    }
}