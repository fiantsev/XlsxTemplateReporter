using ExcelReportCreatorProject.Domain;
using ExcelReportCreatorProject.Domain.Data;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject
{
    public class InjectionContext
    {
        public Marker Marker { get; set; }
        public ResourceObject ResourceObject { get; set; }
        public IWorkbook Workbook { get; set; }
    }
}