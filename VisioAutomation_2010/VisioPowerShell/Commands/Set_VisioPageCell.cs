using System.Collections;
using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageCell)]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true,Position=0)] 
        public Hashtable Hashtable  { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        protected override void ProcessRecord()
        {
            var target_pages = this.Pages ?? new[] { this.Client.Page.Get() };

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;
                var t = new VisioAutomation.Scripting.TargetShapes(pagesheet);

                var dic = CellHashtableToDictionary(this.Hashtable);
                this.Client.ShapeSheet.SetPageCells( t , dic, this.BlastGuards, this.TestCircular);
            }
        }

        public static Dictionary<string, string> CellHashtableToDictionary(Hashtable ht)
        {
            var dic = new Dictionary<string, string>();

            foreach (object key in ht.Keys)
            {
                if (!(key is string))
                {
                    string message =
                        string.Format("Only string values can be keys in the hashtable. Encountered a key of type {0}",
                            key.GetType().FullName);
                    throw new System.ArgumentOutOfRangeException(message);
                }

                string cellname = (string) key;
                var cell_value_o = ht[key];

                if (!(cell_value_o is string))
                {
                    string message =
                        string.Format("Only string values can be values in the hashtable. Encountered a key of type {0}",
                            key.GetType().FullName);
                    throw new System.ArgumentOutOfRangeException(message);

                }
                dic[cellname] = (string) cell_value_o;
            }
            return dic;
        }
    }
}