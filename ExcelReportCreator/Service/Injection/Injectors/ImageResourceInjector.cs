using System;
using System.IO;
using ExcelReportCreatorProject.Domain.ResourceObjects;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Service.Injection.Injectors
{
    public class ImageResourceInjector : IResourceInjector
    {
        public Action<InjectionContext> Inject => (InjectionContext context) =>
        {
            var startMarker = context.MarkerRegion.StartMarker;
            var workbook = context.Workbook;
            var sheet = workbook.GetSheetAt(startMarker.Position.SheetIndex);
            var cell = sheet
                .GetRow(startMarker.Position.RowIndex)
                .GetCell(startMarker.Position.CellIndex);
            var imageResource = (context.ResourceObject as ImageResourceObject);

            //убираем маркер
            cell.SetCellValue("");

            var pictureIndex = workbook.AddPicture(imageResource.Image, PictureType.JPEG);

            var drawing = sheet.CreateDrawingPatriarch();

            //var a = drawing.CreateAnchor(0, 0, 0, 0, startMarker.Position.RowIndex, startMarker.Position.CellIndex, startMarker.Position.RowIndex, startMarker.Position.CellIndex);
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = startMarker.Position.CellIndex;
            anchor.Row1 = startMarker.Position.RowIndex;
            anchor.Col2 = anchor.Col1 + 1;
            anchor.Row2 = anchor.Row1 + 1;
            //anchor.D

            var picture = (NPOI.XSSF.UserModel.XSSFPicture)drawing.CreatePicture(anchor, pictureIndex);

            picture.Resize();
        };

        //private void AddImageToWorkBook(Image img, int colIndex, int rowIndex, XSSFWorkbook workbook, ISheet sheet)
        //{
        //    var ms = new MemoryStream();
        //    img.Save(ms, ImageFormat.Png); 
        //    byte[] data = ms.ToArray();
        //    int pictureIndex = workbook.AddPicture(data, PictureType.PNG);
        //    ICreationHelper helper = workbook.GetCreationHelper(); 
        //    IDrawing drawing = sheet.CreateDrawingPatriarch(); 
        //    IClientAnchor anchor = helper.CreateClientAnchor(); 
        //    anchor.AnchorType = AnchorType.MoveDontResize; 
        //    anchor.Col1 = colIndex;//0 index based column 
        //    anchor.Row1 = rowIndex;//0 index based row 
        //    IPicture picture = drawing.CreatePicture(anchor, pictureIndex); 
        //    picture.Resize(); 
        //}


        public static byte[] ConvertImageToByteArray(string imagePath)
        {
            byte[] imageByteArray = null;
            FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
            using (BinaryReader reader = new BinaryReader(fileStream))
            {
                imageByteArray = new byte[reader.BaseStream.Length];
                for (int i = 0; i < reader.BaseStream.Length; i++)
                    imageByteArray[i] = reader.ReadByte();
            }
            return imageByteArray;
        }
    }
}
