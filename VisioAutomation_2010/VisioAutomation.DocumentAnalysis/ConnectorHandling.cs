namespace VisioAutomation.DocumentAnalysis
{
    public class ConnectorHandling
    {
        public ArrowHandling AR = ArrowHandling.UseConnectorArrows;
        public NoArrowsHandling ADR = NoArrowsHandling.TreatAsBidirectional;
    }

    public enum ArrowHandling
    {
        UseConnectionOrder,
        UseConnectorArrows
    }

    public enum NoArrowsHandling
    {
        Exclude,
        TreatAsBidirectional,
    }
}