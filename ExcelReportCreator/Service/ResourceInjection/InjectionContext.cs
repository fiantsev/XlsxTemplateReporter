using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Domain.Markers;
using ClosedXML.Excel;

namespace ExcelReportCreatorProject.Service.ResourceInjection
{
    public class InjectionContext
    {
        public MarkerRange MarkerRange { get; set; }
        public ResourceObject ResourceObject { get; set; }
        public IXLWorkbook Workbook { get; set; }
    }
}