using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Domain.Markers;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Service.Injection
{
    public class InjectionContext
    {
        public MarkerRegion MarkerRegion { get; set; }
        public ResourceObject ResourceObject { get; set; }
        public IWorkbook Workbook { get; set; }
    }
}