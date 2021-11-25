// See https://aka.ms/new-console-template for more information
using VisioAutomation.Geometry;
using VisioScripting;

Console.WriteLine("Hello, World!");


var app = new Microsoft.Office.Interop.Visio.Application();
var ss = new VisioScripting.Client(app);

var r = new Rectangle(0, 0, 4, 5);

ss.Document.NewDocument();
var tp = TargetPage.Auto;
ss.Draw.DrawRectangle(tp,r);