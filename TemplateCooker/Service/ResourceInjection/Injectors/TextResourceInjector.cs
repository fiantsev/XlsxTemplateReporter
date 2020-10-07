using System;
using TemplateCooker.Domain.ResourceObjects;
using TemplateCooker.Service.Utils;

namespace TemplateCooker.Service.ResourceInjection.Injectors
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