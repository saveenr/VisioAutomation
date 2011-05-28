using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomationSamples
{
    public class Program
    {
        private static void Main(string[] args)
        {
            bool debug = false;

            if (!debug)
            {
                EffectsSamples.SoftShadow();
                EffectsSamples.GradientTransparencies();
                PlaygroundSamples.DrawAllGradients();
                PlaygroundSamples.Spirograph();
                SmartShapeSamples.ProgressBar();
                CustomPropertySamples.SetCustomProperties();
                PathAnalysisSamples.PathAnalysis();
                SimpleGeometrySamples.BezierCircle();
                SimpleGeometrySamples.BezierEllipse();
                SimpleGeometrySamples.BezierSimple();
                SimpleGeometrySamples.NURBS2();
                SimpleGeometrySamples.NURBS3();
                InfoGraphicSamples.BarChart();
                InfoGraphicSamples.PieChart();
                TextSamples.TextMarkup2();
                TextSamples.TextSizing();
                TextSamples.NonRotatingText();
                TextSamples.TextFields();
                TextSamples.TextMarkup1();
                TextSamples.FontChart();
                LayoutSamples.BoxHierarchy();
                LayoutSamples.BoxHeirarchy_FontGlyphComparision();
                LayoutSamples.MSAGL();
                StencilSamples.DrawGridOfMasters();
                ConnectorSamples.ConnectorsToBack(); // TODO: make connector style a simple direct line           
                ColorSample.ColorGrid();
                // creates new docs
                SpecialDocumentSamples.OrgChart();
            }
            else
            {
                LayoutSamples.BoxHeirarchy_FontGlyphComparision();
            }
        }
    }
}