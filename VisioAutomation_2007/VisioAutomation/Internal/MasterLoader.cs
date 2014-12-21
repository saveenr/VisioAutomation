using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Internal
{
    internal class MasterLoader
    {
        private readonly Dictionary<string, MasterRef> master_ref_dic;

        public class MasterRef
        {
            public string StencilName { get; private set; }
            public string MasterName { get; private set; }
            public IVisio.Master VisioMaster { get; set; }

            public MasterRef(string mastername, string stencilname)
            {
                this.MasterName = mastername;
                this.StencilName = stencilname;
                this.VisioMaster = null;
            }
        }
        
        public MasterLoader()
        {
            this.master_ref_dic = new Dictionary<string, MasterRef>();
        }

        public void Add(string mastername, string stencilname)
        {
            var item = new MasterRef(mastername,stencilname);
            string key = this.getkey(mastername, stencilname);
            this.master_ref_dic[key] = item;
        }

        private string getkey(string mastername, string stencilname)
        {
            return mastername + "+" + stencilname;
        }

        public MasterRef Get(string mastername, string stencilname)
        {
            string key = this.getkey(mastername, stencilname);
            return this.master_ref_dic[key];
        }

        public bool Contains(string mastername, string stencilname)
        {
            string key = this.getkey(mastername, stencilname);
            return this.master_ref_dic.ContainsKey(key);
        }

        public void Resolve(IVisio.Documents docs)
        {
            // first get the unique stencils (ignoring case)
            var comparer = System.StringComparer.CurrentCultureIgnoreCase;
            var unique_stencils = new HashSet<string>(comparer);
            foreach (var master_ref in this.master_ref_dic.Values)
            {
                unique_stencils.Add(master_ref.StencilName);
            }

            // for each unique stencil, load the stencil doc
            var name_to_stencildoc = new Dictionary<string, IVisio.Document>(comparer);
            foreach (var stencil in unique_stencils.Where(s=>s!=null))
            {
                // If a stencil was stencified open the stencil if needed
                var stencil_doc = docs.OpenStencil(stencil);
                if (stencil_doc == null)
                {
                    string msg = string.Format("Failed to open stencil \"{0}\"", stencil);
                    throw new AutomationException(msg);
                }

                name_to_stencildoc[stencil] = stencil_doc;
            }

            // identify real master objects for all deferred shapes
            foreach (var master_ref in this.master_ref_dic.Values)
            {
                if (master_ref.VisioMaster == null)
                {
                    if (master_ref.StencilName != null)
                    {
                        // The stencil doc was specified so try to find the master in that stencil doc
                        var stencildoc = name_to_stencildoc[master_ref.StencilName];
                        var stencilmasters = stencildoc.Masters;

                        var master_object = this.TryGetMaster(stencilmasters, master_ref.MasterName);
                        if (master_object == null)
                        {
                            string msg = string.Format("No such master \"{0}\" in stencil \"{1}\"",
                                                       master_ref.MasterName, master_ref.StencilName);
                            throw new AutomationException(msg);
                        }
                        master_ref.VisioMaster = master_object;                        
                    }
                    else
                    {
                        // the stencil doc was not specified so try to find the master int the current doc
                        var app = docs.Application;
                        var stencildoc = app.ActiveDocument;
                        var stencilmasters = stencildoc.Masters;

                        var master_object = this.TryGetMaster(stencilmasters, master_ref.MasterName);
                        if (master_object == null)
                        {
                            string msg = string.Format("No such master \"{0}\" in Active Document \"{1}\"",
                                                       master_ref.MasterName, stencildoc.Name);
                            throw new AutomationException(msg);
                        }
                        master_ref.VisioMaster = master_object;
                    }
                }
            }
        }

        private IVisio.Master TryGetMaster(IVisio.Masters masters, string name)
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
    }
}