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

        public IDictionary<IVisio.Shape, Dictionary<string,VACUSTPROP.CustomPropertyCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, VACUSTPROP.CustomPropertyCells>>();
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this.Client.Application.Get();
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

        public IList<bool> Contains(IList<IVisio.Shape> target_shapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<bool>();
            }

            var results = this.Client.Selection.GetShapes().Select(s => VACUSTPROP.CustomPropertyHelper.Contains(s, name)).ToList();

            return results;
        }

        public void Delete(IList<IVisio.Shape> target_shapes, string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Delete Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Delete(shape, name);
                }
            }
        }

        public void Set(IList<IVisio.Shape> target_shapes, string name, VACUSTPROP.CustomPropertyCells customprop)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            if (customprop == null)
            {
                throw new System.ArgumentNullException(nameof(customprop));
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Set Custom Property"))
            {
                foreach (var shape in shapes)
                {
                    VACUSTPROP.CustomPropertyHelper.Set(shape, name, customprop);
                }
            }
        }

        public IEnumerable<IVisio.Shape> EnumerateAndSelect(IEnumerable<IVisio.Shape> shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            foreach (var shape in shapes)
            {
                this.Client.Selection.None();
                this.Client.Selection.Select(shape);
                yield return shape;
            }
        }
    }
}