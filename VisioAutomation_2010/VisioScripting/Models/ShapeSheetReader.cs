using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

using VisioAutomation.Extensions;

namespace VisioScripting.Models
{
    public class ShapeSheetReader
    {
        public Client Client;
        public IVisio.Page Page;
        public List<VisioAutomation.Core.SidSrc> SidSrcs;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Page = page;
            this.SidSrcs = new List<VisioAutomation.Core.SidSrc>();
        }

        public void AddCell(short id, VisioAutomation.Core.Src src)
        {
            var sidsrc = new VisioAutomation.Core.SidSrc(id, src);
            this.SidSrcs.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var stream = VisioAutomation.ShapeSheet.Streams.StreamArray.FromSidSrc(this.SidSrcs);
            var formulas = this.Page.GetFormulasU(stream);
            return formulas;
        }

        public string[] GetResults()
        {
            const object [] unitcodes = null;
            var stream = VisioAutomation.ShapeSheet.Streams.StreamArray.FromSidSrc(this.SidSrcs);
            var formulas = this.Page.GetResults<string>( stream, unitcodes);
            return formulas;
        }
    }
}