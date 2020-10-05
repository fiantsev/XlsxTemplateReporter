using ExcelReportCreatorProject.Domain.Markers;
using ExcelReportCreatorProject.Service.ResourceInjection;
using ExcelReportCreatorProject.Service.ResourceObjectProvision;

namespace ExcelReportCreatorProject.Service.Creation
{
    public class DocumentInjectorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IResourceObjectProvider ResourceObjectProvider { get; set; }
        public MarkerOptions MarkerOptions { get; set; }
    }
}