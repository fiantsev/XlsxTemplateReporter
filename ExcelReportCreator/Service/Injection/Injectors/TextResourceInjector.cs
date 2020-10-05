using System;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using ExcelReportCreatorProject.Service.Utils;

namespace ExcelReportCreatorProject.Service.Injection.Injectors
{
    public class TextResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => context =>
        {
            var markerPosition = context.MarkerRange.StartMarker.Position;

            var cell = context.Workbook
                .Worksheet(markerPosition.SheetIndex)
                .Row(markerPosition.RowIndex)
                .Cell(markerPosition.CellIndex);

            var text = (context.ResourceObject as TextResourceObject).Text;

            CellUtils.SetDynamicCellValue(cell, text);
        };
    }
}