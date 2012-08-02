﻿using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioCSharpSamples
{

    internal class Program
    {
        private static void Main(string[] args)
        {
            var app = new IVisio.Application();
            var doc = app.Documents.Add("");
            Samples.Shape_Format_Paragraph_Range(doc);
        }
    }

}