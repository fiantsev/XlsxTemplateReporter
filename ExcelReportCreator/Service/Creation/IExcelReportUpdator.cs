using ClosedXML.Excel;

namespace ExcelReportCreatorProject
{
    public interface IExcelReportUpdator
    {
        void Update(IXLWorkbook workbook);
    }
}