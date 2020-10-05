
using ClosedXML.Excel;

namespace ExcelReportCreatorProject
{
    public interface IExcelReportCreator
    {
        IXLWorkbook Create(IXLWorkbook workbook);
    }
}