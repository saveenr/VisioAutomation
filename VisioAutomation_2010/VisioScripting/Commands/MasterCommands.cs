using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class MasterCommands : CommandSet
    {
        internal MasterCommands(Client client) :
            base(client)
        {

        }

        public List<IVisio.Master> GetMasters(TargetDocument targetdoc)
        {
            var doc_masters = targetdoc.Document.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public IVisio.Master GetMaster(TargetDocument targetdoc, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var masters = targetdoc.Document.Masters;
            IVisio.Master masterobj = this._try_get_master(masters, name);

            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", name, targetdoc.Document.Name);
                throw new VisioOperationException(msg);
            }

            return masterobj;
        }


        public List<IVisio.Master> FindMasters(TargetDocument targetdoc, string name)
        {
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all masters
                var masters = targetdoc.Document.Masters.ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = targetdoc.Document.Masters.ToEnumerable();
                var masters3 = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(masters2, new[] { name }, p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                return masters3;
            }
        }

        private IVisio.Master _try_get_master(IVisio.Masters masters, string name)
        {
            try
            {
                var masterobj = masters.ItemU[name];
                return masterobj;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        public IVisio.Shape DropMaster(TargetPage targetpage, IVisio.Master master, VisioAutomation.Core.Point p)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            var shape = targetpage.Page.Drop(master, p.X, p.Y);
            return shape;
        }

        public short[] DropMasters(
            TargetPage targetpage,
            IList<IVisio.Master> masters, 
            IList<VisioAutomation.Core.Point> points)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            var page = targetpage.Page;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }
    }
}