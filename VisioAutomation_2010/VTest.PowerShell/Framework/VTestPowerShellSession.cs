using SMA = System.Management.Automation;

namespace VTest.PowerShell.Framework
{
    public class VTestPowerShellSession : System.IDisposable 
    {
        protected SMA.PowerShell _powershell;
        protected SMA.Runspaces.InitialSessionState _sessionstate;
        protected SMA.Runspaces.Runspace _runspace;
        protected SMA.RunspaceInvoke _invoker;

        public VTestPowerShellSession()
        {
            this._sessionstate = SMA.Runspaces.InitialSessionState.CreateDefault();
            this._runspace = SMA.Runspaces.RunspaceFactory.CreateRunspace(this._sessionstate);
            this._runspace.Open();
            this._powershell = SMA.PowerShell.Create();
            this._powershell.Runspace = this._runspace;
            this._invoker = new SMA.RunspaceInvoke(this._runspace);
        }

        public void CleanUp()
        {
            // Make sure we cleanup everything
            if (this._powershell != null)
            {
                this._powershell.Dispose();
                this._powershell = null;
            }
            if (this._invoker != null)
            {
                this._invoker.Dispose();
                this._invoker = null;
            }
            if (this._runspace != null)
            {
                this._runspace.Close();
                this._runspace = null;
            }

            this._sessionstate = null;
        }

        public void Dispose()
        {
            this.CleanUp();
        }
    }
}