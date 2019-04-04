using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetObject<T> where  T: class
    {
        public readonly T Item;
        public readonly bool UseContext;

        public TargetObject()
        {
            this.Item = null;
            this.UseContext = true;
        }
        
        public TargetObject(T item)
        {
            this.Item = item;
            this.UseContext = this.Item == null;
        }
        public TargetObject(T item, bool isresolved)
        {
            this.Item = item;
            this.UseContext = !isresolved;
        }

        public bool IsResolved => !this.UseContext;
    }

    public class TargetDocument: TargetObject<IVisio.Document>
    {
        public TargetDocument() :base()
        {
        }

        public TargetDocument(IVisio.Document doc) : base(doc)
        {
        }

        public TargetDocument(IVisio.Document doc, bool isresolved) : base(doc, isresolved)
        {
        }

        public TargetDocument Resolve(VisioScripting.Client client)
        {
            if (this.IsResolved)
            {
                return this;
            }
            var cmdtarget = client.GetCommandTarget(
                Commands.CommandTargetRequirementFlags.RequireApplication |
                Commands.CommandTargetRequirementFlags.RequireActiveDocument |
                Commands.CommandTargetRequirementFlags.RequirePage);

            // It doesn't matter if there is an active document or not
            // at this point it is considered resolved
            return new TargetDocument(cmdtarget.ActiveDocument, true);
        }
    }
}