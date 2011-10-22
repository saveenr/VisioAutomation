using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IVisio=Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA=VisioAutomation;

namespace VisioExportPagesToDocs
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                System.Console.WriteLine("Syntax is: VisioExportPagesToDocs <filename.vsd>");
                System.Environment.Exit(0);
            }
            string filename = args[0];
            string absfilename = System.IO.Path.GetFullPath(filename);

            if (!System.IO.File.Exists(absfilename))
            {
                System.Console.WriteLine("File does not exist: \"{0}\"",absfilename);
                System.Environment.Exit(0);
            }

            var visioapp = new IVisio.Application();
            var docs = visioapp.Documents;
            IVisio.Document doc=null;
            try
            {
                doc = docs.Open(absfilename);

            }
            catch (System.Runtime.InteropServices.COMException comexc)
            {
                System.Console.WriteLine("Failed to open file: {0}", comexc.Message);
                System.Environment.Exit(0);
            }

            var pages = doc.Pages;
            var _pages = pages.AsEnumerable().ToList();

            string destpath = System.IO.Path.GetDirectoryName(absfilename);
            string basename = System.IO.Path.GetFileNameWithoutExtension(absfilename);
            string ext = System.IO.Path.GetExtension(absfilename);

            int pageindex = 1;
            foreach (var page in _pages)
            {
                string pagename = page.Name;
                var newdoc = docs.Add("");
                var newpage = newdoc.Pages[1];
                VA.Pages.PageHelper.DuplicateToDocument(page,newdoc,newpage,pagename,true);
                string destname = System.IO.Path.Combine(destpath, basename + "_" + pageindex.ToString() + "_" + pagename + ext);
                if (System.IO.File.Exists(destname))
                {
                    System.Console.WriteLine("Output file already exists. Skipping. File = \"{0}\"",destname);
                }
                newdoc.SaveAs(destname);
                newdoc.Close(true);
                pageindex++;
            }
            doc.Close(true);
            visioapp.Quit(true);
        }
    }
}
