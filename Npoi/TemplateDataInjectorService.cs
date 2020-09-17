using System;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;

namespace Npoi
{
    public class TemplateDataInjectorService
    {
        public void InjectData(
            string filePath,
            Dictionary<string, WidgetData> dataSet
        )
        {
            var workbook = new XSSFWorkbook(filePath);
            Console.WriteLine(workbook.NumberOfSheets);
            Console.WriteLine(workbook.GetSheet(""));
            Console.WriteLine(workbook.GetSheet("Sheet1"));
            Console.WriteLine(workbook.GetSheet("Лист"));
            Console.WriteLine(workbook.GetSheetAt(0));
        }
    }
}
