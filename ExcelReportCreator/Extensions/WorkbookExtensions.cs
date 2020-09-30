using System.Collections.Generic;
using ExcelReportCreatorProject.Domain.Npoi;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Extensions
{
    public static class WorkbookExtensions
    {
        public static IEnumerable<ISheet> EnumerateSheets(this IWorkbook workbook)
        {
            return new SheetCollection(workbook);
        }
    }
}
