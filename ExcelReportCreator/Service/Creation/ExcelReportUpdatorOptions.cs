using ExcelReportCreatorProject.Service.Extraction;
using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;

namespace ExcelReportCreatorProject.Service.Creation
{
    public class ExcelReportUpdatorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IResourceObjectProvider ResourceObjectProvider { get; set; }
        public IMarkerExtractor MarkerExtractor { get; set; }
        public FormulaEvaluationOptions FormulaEvaluationOptions { get; set; }
    }
}