using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAnalyzeFontUsage
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename = @"D:\\alsbd008.vsd";

            var app = new IVisio.Application();

            app.Documents.Add(filename);

            var shapes = app.ActivePage.Shapes.AsEnumerable().ToList();
            var query = new VA.ShapeSheet.Query.SectionQuery(IVisio.VisSectionIndices.visSectionCharacter);
            query.AddColumn(IVisio.VisCellIndices.visCharacterFont);
            query.AddColumn(IVisio.VisCellIndices.visCharacterSize);
            var shapeids = shapes.Select(s => s.ID).ToList();
            var results = query.GetResults<double>(app.ActivePage, shapeids);

            foreach (var g in results.Groups)
            {
                foreach (int r in g.RowIndices)
                {
                    double font = results[r, 0];
                    Console.WriteLine("Font: {0}", font);
                }
            }
        }
    }
}
