using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Domain.Markers;
using ClosedXML.Excel;

namespace ExcelReportCreatorProject.Service.Injection
{
    public class InjectionContext
    {
        public MarkerRegion MarkerRegion { get; set; }
        public ResourceObject ResourceObject { get; set; }
        public IXLWorkbook Workbook { get; set; }
    }
}