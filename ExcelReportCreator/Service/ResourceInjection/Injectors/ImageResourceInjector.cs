using System;
using System.IO;
using ClosedXML.Excel;
using ExcelReportCreatorProject.Domain.ResourceObjects;

namespace ExcelReportCreatorProject.Service.ResourceInjection.Injectors
{
    public class ImageResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext context) =>
        {
            var startMarker = context.MarkerRange.StartMarker;
            var workbook = context.Workbook;
            var sheet = workbook.Worksheet(startMarker.Position.SheetIndex);
            var cell = sheet
                .Row(startMarker.Position.RowIndex)
                .Cell(startMarker.Position.CellIndex);
            var imageResource = (context.ResourceObject as ImageResourceObject);

            //убираем маркер
            cell.Clear(XLClearOptions.Contents);

            using (var imageStream = new MemoryStream(imageResource.Image))
            {
                var image = sheet.AddPicture(imageStream)
                  .MoveTo(cell)
                  .Scale(1);
            }
        };
    }
}