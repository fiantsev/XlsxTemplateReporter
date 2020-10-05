using System;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Service.Utils;

namespace ExcelReportCreatorProject.Service.Injection.Injectors
{
    public class TableResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            var table = (injectionContext.ResourceObject as TableResourceObject).Table;

            var rowCount = table.Count;
            var columnCount = table.Count == 0
                ? 0
                : table[0].Count;

            var insertionStartRowIndex = injectionContext.MarkerRange.StartMarker.Position.RowIndex;
            var insertionStartCellIndex = injectionContext.MarkerRange.StartMarker.Position.CellIndex;

            var sheet = injectionContext.Workbook.Worksheet(injectionContext.MarkerRange.StartMarker.Position.SheetIndex);

            for (var dataRowIndex = 0; dataRowIndex < rowCount; ++dataRowIndex)
            {
                var dataRow = table[dataRowIndex];
                var currentRowIndex = insertionStartRowIndex + dataRowIndex;
                var currentRow = sheet.Row(currentRowIndex);

                //if (currentRow == null)
                //    currentRow = sheet.Rows(). . CreateRow(currentRowIndex);

                for (var dataColIndex = 0; dataColIndex < columnCount; ++dataColIndex)
                {
                    var dataValue = dataRow[dataColIndex];
                    var currentCellIndex = insertionStartCellIndex + dataColIndex;
                    var currentCell = currentRow.Cell(currentCellIndex);

                    //if (currentCell == null)
                    //    currentCell = currentRow.CreateCell(currentCellIndex);

                    CellUtils.SetDynamicCellValue(currentCell, dataValue);
                }
            }
        };
    }
}
