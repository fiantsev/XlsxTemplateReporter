using System;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class TableResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            var markerPosition = injectionContext.MarkerRange.StartMarker.Position;
            var table = (injectionContext.ResourceObject as TableResourceObject).Table;

            var rowCount = table.Count;
            var columnCount = table.Count == 0
                ? 0
                : table[0].Count;

            var sheet = injectionContext.Workbook.Worksheet(markerPosition.SheetIndex);

            var topLeftCell = sheet.Cell(markerPosition.RowIndex, markerPosition.CellIndex);
            var mergedRowsEnumerator = CellUtils.EnumerateMergedRows(topLeftCell).GetEnumerator();

            table.ForEach(dataRow =>
            {
                mergedRowsEnumerator.MoveNext();
                var excelRow = mergedRowsEnumerator.Current;

                var firstCellOfRow = sheet.Cell(excelRow.FirstCell().Address.RowNumber, topLeftCell.Address.ColumnNumber);
                var mergedCellsEnumerator = CellUtils.EnumerateMergedCells(firstCellOfRow).GetEnumerator();

                dataRow.ForEach(dataValue =>
                {
                    mergedCellsEnumerator.MoveNext();
                    CellUtils.SetDynamicCellValue(mergedCellsEnumerator.Current, dataValue);
                });
            });
        };
    }
}