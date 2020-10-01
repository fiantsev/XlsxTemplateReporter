using ExcelReportCreatorProject.Domain.Markers.ExtractorOptions;
using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;

namespace ExcelReportCreatorProject.Service.Creator
{
    public class ExcelReportCreatorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IResourceObjectProvider ResourceObjectProvider { get; set; }
        public MarkerExtractionOptions MarkerExtractionOptions { get; set; }
        public FormulaEvaluationOptions FormulaEvaluationOptions { get; set; }
    }
}