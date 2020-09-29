using ExcelReportCreatorProject.Domain;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExcelReportCreatorProject.Service
{
    public interface IExcelReportParser
    {
        IEnumerable<Marker> GetMarkers(IWorkbook workbook);
        IEnumerable<Marker> GetMarkers(ISheet sheet);
    }
}