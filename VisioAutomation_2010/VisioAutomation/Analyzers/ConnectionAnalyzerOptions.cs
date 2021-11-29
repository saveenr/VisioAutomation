namespace VisioAutomation.Analyzers
{
    public class ConnectionAnalyzerOptions
    {
        public EdgeDirectionSource EdgeDirectionSource = EdgeDirectionSource.UseConnectorArrows;
        public EdgeNoArrowsHandling EdgeNoArrowsHandling = EdgeNoArrowsHandling.IncludeEdgesForBothDirections;
    }
}