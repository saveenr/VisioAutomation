using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Internal
{
    public class MasterLoader
    {
        public class MasterRef
        {
            public string StencilName;
            public string MasterName;
            public IVisio.Master VisioMaster;
        }

        private Dictionary<string, MasterRef> dic;

        public MasterLoader()
        {
            this.dic = new Dictionary<string, MasterRef>();
        }

        public void Add(string mastername, string stencilname)
        {
            var item = new MasterRef();
            item.MasterName = mastername;
            item.StencilName = stencilname;
            item.VisioMaster = null;

            string key = this.getkey(mastername, stencilname);
            this.dic[key] = item;
        }

        private string getkey(string mastername, string stencilname)
        {
            return mastername + "+" + stencilname;
        }

        public MasterRef Get(string mastername, string stencilname)
        {
            string key = this.getkey(mastername, stencilname);
            return this.dic[key];
        }

        public bool Contains(string mastername, string stencilname)
        {
            string key = this.getkey(mastername, stencilname);
            return this.dic.ContainsKey(key);
        }


        public void Resolve(IVisio.Documents docs)
        {
            var unique_stencils = new HashSet<string>();
            foreach (var kv in this.dic)
            {
                string mr = kv.Value.StencilName;
                unique_stencils.Add(mr);
            }
            var name_to_stencildoc = new Dictionary<string, IVisio.Document>();
            foreach (var stencil in unique_stencils)
            {
                try
                {
                    var stencil_doc = docs.OpenStencil(stencil);
                    name_to_stencildoc[stencil] = stencil_doc;
                }
                catch (Exception)
                {
                    string msg = string.Format("No such Stencil \"{0}\"", stencil);
                    throw new AutomationException(msg);
                }
            }

            // identify real master objects for all deferred shapes
            foreach (var mr in this.dic.Values)
            {
                if (mr.VisioMaster == null)
                {
                    var stencildoc = name_to_stencildoc[mr.StencilName];
                    var stencilmasters = stencildoc.Masters;

                    try
                    {
                        var master_object = stencilmasters[mr.MasterName];
                        mr.VisioMaster = master_object;
                    }
                    catch (Exception)
                    {
                        string msg = string.Format("No such Master \"{0}\" in Stencil \"{1}\"",
                                                   mr.MasterName, mr.StencilName);
                        throw new AutomationException(msg);
                    }
                }
            }
        }

    }
}