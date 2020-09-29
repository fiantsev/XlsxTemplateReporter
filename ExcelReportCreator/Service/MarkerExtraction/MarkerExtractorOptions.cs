using ExcelReportCreatorProject.Domain;

namespace ExcelReportCreatorProject.Service.MarkerExtraction
{
    public class MarkerExtractorOptions
    {
        public MarkerOptions MarkerOptions { get; set; }
        public MarkerExtractionTechnique Technique { get; set; }
    }
}