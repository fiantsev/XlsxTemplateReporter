using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExcelReportCreatorProject.Service.Injection
{
    public static class InjectionExtensions
    {
        public static void InjectListOfList(this InjectionContext injectionContext, List<List<string>> dataSet)
        {
            var dataSetRowCount = dataSet.Count;
            var dataSetColCount = dataSet.Count == 0
                ? 0
                : dataSet[0].Count;

            var insertionStartRowIndex = injectionContext.MarkerRegion.StartMarker.Position.RowIndex;
            var insertionStartCellIndex = injectionContext.MarkerRegion.StartMarker.Position.CellIndex;

            var sheet = injectionContext.Workbook.GetSheetAt(injectionContext.MarkerRegion.StartMarker.Position.SheetIndex);

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
        }
    }
}