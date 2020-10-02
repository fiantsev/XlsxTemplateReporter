using ExcelReportCreatorProject.Domain.ResourceObjects;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelReportCreatorProject.Service.Injection.Injectors
{
    public class ImageResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext context) =>
        {
            var startMarker = context.MarkerRegion.StartMarker;
            var workbook = (XSSFWorkbook)context.Workbook;
            var sheet = workbook.GetSheetAt(startMarker.Position.SheetIndex);
            var cell = sheet
                .GetRow(startMarker.Position.RowIndex)
                .GetCell(startMarker.Position.CellIndex);
            var imageResource = (context.ResourceObject as ImageResourceObject);

            //убираем маркер
            cell.SetCellValue("");

            var drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            var addDimensionedImage = new AddDimensionedImage();

            var imageWidthInMm = AddDimensionedImage.ConvertImageUnits.widthUnits2Millimetres(
                AddDimensionedImage.ConvertImageUnits.pixel2WidthUnits(884)
            );
            var imageHeightInMm = AddDimensionedImage.ConvertImageUnits.widthUnits2Millimetres(
                AddDimensionedImage.ConvertImageUnits.pixel2WidthUnits(2392)
            );

            addDimensionedImage.addImageToSheet(
                (startMarker.Position.RowIndex, startMarker.Position.CellIndex),
                sheet,
                drawing,
                (imageResource.Image, PictureType.PNG),
                imageWidthInMm,
                imageHeightInMm,
                AddDimensionedImage.OVERLAY_ROW_AND_COLUMN
            );

            //var pictureIndex = workbook.AddPicture(imageResource.Image, PictureType.PNG);

            //var anchor = (XSSFClientAnchor)((XSSFCreationHelper)workbook.GetCreationHelper()).CreateClientAnchor();
            //anchor.Col1 = startMarker.Position.CellIndex;
            //anchor.Row1 = startMarker.Position.RowIndex;

            //var picture = (XSSFPicture)drawing.CreatePicture(anchor, pictureIndex);

            //picture.Resize(1);
        };
    }
}