using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject
{
    public interface IExcelReportCreator
    {
        IWorkbook Create(IWorkbook workbook);
    }
}