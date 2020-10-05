using ClosedXML.Excel;

namespace ExcelReportCreatorProject.Service.Utils
{
    public class SheetUtils
    {
        public static int SheetIndex(IXLWorksheet sheet)
        {
            return sheet.Position;
        }
    }
}