using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using TemplateCooker.Domain.Injections;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.ResourceInjection.Injectors
{
    public class TableResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext injectionContext) =>
        {
            ShiftLayout(injectionContext);
            InsertTable(injectionContext);
        };

        private void ShiftLayout(InjectionContext injectionContext)
        {
            var markerRange = injectionContext.MarkerRange;
            var injection = (injectionContext.Injection as TableInjection);
            var table = injection.Resource.Object;

            switch (injection.LayoutShift)
            {
                case LayoutShiftType.None:
                    return;
                case LayoutShiftType.MoveRows:
                    injectionContext.Workbook.Worksheet(markerRange.StartMarker.Position.SheetIndex)
                        .Row(markerRange.EndMarker.Position.RowIndex)
                        .InsertRowsBelow(table.Count - 1); //-1 потому что ячейка в которой находиться маркер предоставляет одну ячейку под пространство вставки
                    return;
                case LayoutShiftType.MoveCells:
                    throw new Exception("Unsupported");
                default:
                    throw new Exception($"Unhandled case: {nameof(injection.LayoutShift)}={injection.LayoutShift.ToString()}");
            }
        }

        private void InsertTable(InjectionContext injectionContext)
        {
            var markerPosition = injectionContext.MarkerRange.StartMarker.Position;
            var table = (injectionContext.Injection as TableInjection).Resource.Object;

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
        }
    }
}