using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;

namespace Test
{
    public class TemplateParser
    {
        public static void PrintFullInfo(XSSFWorkbook workbook)
        {
            for(var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; ++sheetIndex)
            {
                var sheet = (XSSFSheet)workbook.GetSheetAt(sheetIndex);
                var tables = sheet.GetTables();
                Console.WriteLine($"|sheet SheetName={sheet.SheetName} FirstRowNum={sheet.FirstRowNum} LastRowNum={sheet.LastRowNum} PhysicalNumberOfRows={sheet.PhysicalNumberOfRows} TablesCount={tables.Count} PivotTablesCount={sheet.GetPivotTables().Count}|");

                foreach(var table in tables)
                    Console.WriteLine($"  /table Name={table.DisplayName} ref={table.GetCTTable().@ref} StartCellReference={table.StartCellReference} EndCellReference={table.EndCellReference}/");

                //var names = workbook.NumberOfNames
                for(var nameIndex = 0; nameIndex < workbook.NumberOfNames; ++nameIndex)
                {
                    var name = workbook.GetNameAt(nameIndex);
                    Console.WriteLine($"  /name ToString={name.ToString()} NameName={name.NameName} /");
                }

                for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
                {
                    var row = (XSSFRow)sheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        Console.WriteLine($"  [row {rowIndex} isEmpty]");
                        continue;
                    }
                    Console.WriteLine($"  [row {rowIndex} FirstCellNum={row.FirstCellNum} LastCellNum={row.LastCellNum} PhysicalNumberOfCells={row.PhysicalNumberOfCells}]");

                    for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                    {
                        var cell = (XSSFCell)row.GetCell(cellIndex);
                        if(cell == null)
                        {
                            Console.WriteLine($"    <cell {rowIndex}.{cellIndex} isEmpty>");
                            continue;
                        }
                        Console.WriteLine($"    <cell {cell.RowIndex}.{cell.ColumnIndex} type={cell.CellType} value={cell.ToString()} Reference={cell.GetReference()}> ");
                    }
                    Console.WriteLine();
                }
            }
            Console.WriteLine();
        }
    }
}
