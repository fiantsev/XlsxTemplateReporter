using ClosedXML.Excel;

namespace ExcelReportCreatorProject
{
    public interface IDocumentInjector
    {
        void Inject(IXLWorkbook workbook);
    }
}