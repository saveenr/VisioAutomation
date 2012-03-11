using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation
{
    public class UndoScope : System.IDisposable
    {
        private static int scope_count;
        private System.DateTimeOffset time_closed;

        public int ScopeID { get; private set; }
        public IVisio.Application Application { get; private set; }
        public int SequenceNumber { get; private set; }
        public string Name { get; private set; }
        public bool IsOpen { get; private set; }
        public DateTimeOffset OpenedOn { get; private set; }
        public bool Commit { get; set; }

        internal UndoScope(IVisio.Application app, string name)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException("app");
            }

            if (string.IsNullOrEmpty(name))
            {
                throw new System.ArgumentException("name");
            }

            this.Application = app;
            this.Name = name;
            this.ScopeID = this.Application.BeginUndoScope(name);
            this.Commit = true;
            this.OpenedOn = System.DateTimeOffset.UtcNow;
            this.SequenceNumber = scope_count;
            scope_count++;
            this.IsOpen = true;
        }

        /// <summary>
        /// When the scope was closed
        /// </summary>
        public System.DateTimeOffset ClosedOn
        {
            get
            {
                if (this.IsOpen)
                {
                    throw new AutomationException("Scope is not closed");
                }

                return time_closed;
            }
        }

        /// <summary>
        /// Ends the scope if the scope is open. If thw scope is already closed it does nothing.
        /// </summary>
        public void EndScope()
        {
            if (this.IsOpen)
            {
                this.Application.EndUndoScope(this.ScopeID, this.Commit );
                this.IsOpen = false;
                this.time_closed = System.DateTimeOffset.UtcNow;
            }
        }

        /// <summary>
        /// Dispose will end the scope if the scope is still open
        /// </summary>
        public void Dispose()
        {
            this.EndScope();
        }

        /// <summary>
        /// A human-readable description of the scop
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string start = this.OpenedOn.ToString();

            string end = "NA";
            if (!IsOpen)
            {
                end = this.ClosedOn.ToString();
            }

            string s = string.Format("UndoScope({0},Begin={1},End={2})", this.Name, start, end);
            return s;
        }
    }
}