using ExcelReportCreatorProject.Service.Injection;
using ExcelReportCreatorProject.Service.MarkerExtraction;
using ExcelReportCreatorProject.Service.ResourceObjectProvider;

namespace ExcelReportCreatorProject.Service.Creator
{
    public class ExcelReportCreatorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IResourceObjectProvider ResourceObjectProvider { get; set; }
        public MarkerExtractorOptions MarkerExtractorOptions { get; set; }
    }
}