using System.Collections.Generic;
using System.Linq;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class CustomPropCommands : CommandSet
    {
        internal CustomPropCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, Dictionary<string,VACUSTPROP.CustomPropertyCells>> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, VACUSTPROP.CustomPropertyCells>>();
            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;

            var list_custom_props = VACUSTPROP.CustomPropertyHelper.Get(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_custom_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(TargetShapes targets, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var shapes = targets.ResolveShapes(this._client);


            if (shapes.Count < 1)
            {
                return new List<bool>();
            }

            var results = this._client.Selection.GetShapes().Select(s => VACUSTPROP.CustomPropertyHelper.Contains(s, name)).ToList();

            return results;
        }

        public void Delete(TargetShapes targets, string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Delete Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(TargetShapes  targets, string name, VACUSTPROP.CustomPropertyCells customprop)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (customprop == null)
            {
                throw new System.ArgumentNullException(nameof(customprop));
            }

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }

        public IEnumerable<IVisio.Shape> EnumerateAndSelect(IEnumerable<IVisio.Shape> shapes)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            foreach (var shape in shapes)
            {
                this._client.Selection.SelectNone();
                this._client.Selection.Select(shape);
                yield return shape;
            }
        }
    }
}