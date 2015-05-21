using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Application
{
    public class UndoScope : System.IDisposable
    {
        private readonly IVisio.Application Application;

        public int ScopeID { get; private set; }
        public string Name { get; private set; }
        public bool Commit { get; set; }

        public UndoScope(IVisio.Application app, string name)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException(nameof(app));
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                string msg = $"{"name"} cannot be null or empty";
                throw new System.ArgumentException(msg,nameof(name));
            }

            this.Application = app;
            this.Name = name;
            this.ScopeID = this.Application.BeginUndoScope(name);
            this.Commit = true;
        }

        /// <summary>
        /// Dispose will end the scope if the scope is still open
        /// </summary>
        public void Dispose()
        {
            this.Application.EndUndoScope(this.ScopeID, this.Commit);
        }

        /// <summary>
        /// A human-readable description of the scop
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string s = $"UndoScope(Name=\"{this.Name}\",ScopeID={this.ScopeID})";
            return s;
        }
    }
}