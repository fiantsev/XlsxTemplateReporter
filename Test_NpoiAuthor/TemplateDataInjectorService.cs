using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Test
{
    public class TemplateDataInjectorService
    {
        public void InjectData(
            XSSFWorkbook workbook,
            Dictionary<string, WidgetData> dataSet
        )
        {
            var markerInfos = Pass1_ExtractMarkers(workbook);
            var dataInjections = markerInfos.Select(x => new DataInjection
            {
                MarkerInfo = x,
                WidgetData = dataSet[x.MarkerId]
            }).ToList();
            Pass2_InjectData(workbook, dataInjections);
        }

        private List<MarkerInfo> Pass1_ExtractMarkers(
            XSSFWorkbook workbook
        )
        {
            var result = new List<MarkerInfo>();
            Enumerable.Range(1, workbook.NumberOfSheets).ToList()
                .ForEach(sheetIndex =>
                {
                    var sheet = (XSSFSheet)workbook.GetSheetAt(sheetIndex);
                    var tableNames = string.Join(",", sheet.GetTables().Select(x => x.DisplayName));
                    //Console.WriteLine($"|sheet SheetName={sheet.SheetName} FirstRowNum={sheet.FirstRowNum} LastRowNum={sheet.LastRowNum} PhysicalNumberOfRows={sheet.PhysicalNumberOfRows} Tables={tableNames} PivotTablesCount={sheet.GetPivotTables().Count}|");

                    for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            //Console.WriteLine($"\t[row {rowIndex} isEmpty]");
                            continue;
                        }
                        //Console.WriteLine($"\t[row {rowIndex} FirstCellNum={row.FirstCellNum} LastCellNum={row.LastCellNum} PhysicalNumberOfCells={row.PhysicalNumberOfCells}]");

                        for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                        {
                            var cell = row.GetCell(cellIndex);
                            if (cell == null)
                                continue;
                            //Console.Write($"\t\t<cell {cell.RowIndex}.{cell.ColumnIndex} type={cell.CellType} value={cell.ToString()}> ");
                            if (IsMarkedCell(cell))
                                result.Add(new MarkerInfo { SheetIndex = sheetIndex, RowIndex = rowIndex, CellIndex = cellIndex, MarkerId = ExtractMarkerValue(cell) });
                        }
                        //Console.WriteLine();
                    }
                });
            //Console.WriteLine();
            return result;
        }

        private bool IsMarkedCell(ICell cell)
        {
            if (cell.CellType == CellType.String)
            {
                var stringCellValue = cell.StringCellValue.Trim();
                if (stringCellValue.Length < 4)
                    return false;
                if (stringCellValue.Substring(0, 2) == "{{" && stringCellValue.Substring(stringCellValue.Length - 2, 2) == "}}")
                    return true;
            }
            return false;
        }

        private string ExtractMarkerValue(ICell cell)
        {
            return cell.StringCellValue.Trim().Substring(2, cell.StringCellValue.Length - 4);
        }



        private void Pass2_InjectData(
            XSSFWorkbook workbook,
            List<DataInjection> dataInjections
        )
        {
            dataInjections.ForEach(dataInjection => InsertCellRangeIntoTable(workbook, dataInjection.MarkerInfo, dataInjection.WidgetData.Values));
        }


        private void InsertCellRange(XSSFWorkbook workbook, MarkerInfo markerInfo, List<List<string>> dataSet)
        {
            var dataSetRowCount = dataSet.Count;
            var dataSetColCount = dataSet.Count == 0
                ? 0
                : dataSet[0].Count;

            var insertionStartRowIndex = markerInfo.RowIndex;
            var insertionStartCellIndex = markerInfo.CellIndex;

            var sheet = workbook.GetSheetAt(markerInfo.SheetIndex);
            for (var dataRowIndex = 0; dataRowIndex < dataSetRowCount; ++dataRowIndex)
            {
                var dataRow = dataSet[dataRowIndex];
                var currentRowIndex = insertionStartRowIndex + dataRowIndex;
                var currentRow = sheet.GetRow(currentRowIndex);

                if (currentRow == null)
                    currentRow = sheet.CreateRow(currentRowIndex);

                for (var dataColIndex = 0; dataColIndex < dataSetColCount; ++dataColIndex)
                {
                    var dataValue = dataRow[dataColIndex];
                    var currentCellIndex = insertionStartCellIndex + dataColIndex;
                    var currentCell = currentRow.GetCell(currentCellIndex);

                    if (currentCell == null)
                        currentCell = currentRow.CreateCell(currentCellIndex);

                    currentCell.SetCellType(CellType.String);
                    currentCell.SetCellValue(dataValue);
                }
            }
        }


        private void InsertCellRangeIntoTable(XSSFWorkbook workbook, MarkerInfo markerInfo, List<List<string>> dataSet)
        {
            var dataSetRowCount = dataSet.Count;
            var dataSetColCount = dataSet.Count == 0
                ? 0
                : dataSet[0].Count;

            var insertionStartRowIndex = markerInfo.RowIndex;
            var insertionStartCellIndex = markerInfo.CellIndex;

            var sheet = (XSSFSheet)workbook.GetSheetAt(markerInfo.SheetIndex);
            var tables = sheet.GetTables();
            sheet.ShiftRows(4, sheet.LastRowNum, 2);
            for (var dataRowIndex = 0; dataRowIndex < dataSetRowCount; ++dataRowIndex)
            {
                var dataRow = dataSet[dataRowIndex];
                var currentRowIndex = insertionStartRowIndex + dataRowIndex;
                var currentRow = sheet.GetRow(currentRowIndex);

                if (currentRow == null)
                    currentRow = sheet.CreateRow(currentRowIndex);

                for (var dataColIndex = 0; dataColIndex < dataSetColCount; ++dataColIndex)
                {
                    var dataValue = dataRow[dataColIndex];
                    var currentCellIndex = insertionStartCellIndex + dataColIndex;
                    var currentCell = currentRow.GetCell(currentCellIndex);

                    if (currentCell == null)
                        currentCell = currentRow.CreateCell(currentCellIndex);

                    currentCell.SetCellType(CellType.Numeric);
                    currentCell.SetCellValue(dataValue);
                }
            }

            tables[0].GetCTTable().@ref = "B2:J6";
        }
    }


    public class MarkerInfo
    {
        public string MarkerId { get; set; }
        public int SheetIndex { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
    }

    public class DataInjection
    {
        public MarkerInfo MarkerInfo { get; set; }
        public WidgetData WidgetData { get; set; }
        public InjectionOptions InjectionOptions { get; set; }
    }

    public class InjectionOptions
    {

    }
}
