using System.Linq;
using ClosedXML.Excel;

namespace ExcelReportCreatorProject.Service.Utils
{
    public class SheetUtils
    {
        public static int SheetIndex(IXLWorksheet sheet)
        {
            return sheet.Workbook.Worksheets.ToList().IndexOf(sheet) + 1;
        }
    }
}