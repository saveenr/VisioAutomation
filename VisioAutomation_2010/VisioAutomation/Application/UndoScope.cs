using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Application
{
    public class UndoScope : System.IDisposable
    {
        private readonly IVisio.Application Application;

        public int ScopeID { get; }
        public string Name { get; }
        public bool Commit { get; set; }

        public UndoScope(IVisio.Application app, string name)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException(nameof(app));
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                string msg = String.Format("{0} cannot be null or empty", "name");
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
            string s = String.Format("UndoScope(Name=\"{0}\",ScopeID={1})", this.Name, this.ScopeID);
            return s;
        }
    }
}