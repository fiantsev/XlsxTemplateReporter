using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject
{
    public interface IExcelReportCreator
    {
        void Create(IWorkbook workbook);
    }
}