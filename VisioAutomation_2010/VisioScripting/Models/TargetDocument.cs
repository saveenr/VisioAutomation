using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
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
            if (!this.IsResolved)
            {
                var cmdtarget = client.GetCommandTarget(
                    Commands.CommandTargetFlags.Application | 
                    Commands.CommandTargetFlags.ActiveDocument |
                    Commands.CommandTargetFlags.ActivePage);

                if (cmdtarget.ActiveDocument!=null)
                {
                    return new TargetDocument(cmdtarget.ActiveDocument);
                }
                else
                {
                    return  new TargetDocument(null,true);
                }
            }
            else
            {
                return this;
            }
        }
    }
}