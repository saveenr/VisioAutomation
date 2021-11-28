namespace VisioAutomation.Analyzers
{
    public class ConnectionAnalyzerOptions
    {
        public DirectionSource DirectionSource = DirectionSource.UseConnectorArrows;
        public EdgeNoArrowsHandling EdgeNoArrowsHandling = EdgeNoArrowsHandling.IncludeEdgesForBothDirections;
    }
}