using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace ExcelReportCreatorProject.Service.Injection
{
    public class AddDimensionedImage
    {

        //  Four constants that determine how - and indeed whether - the rows
        //  and columns an image may overlie should be expanded to accommodate that
        //  image.
        //  Passing EXPAND_ROW will result in the height of a row being increased
        //  to accommodate the image if it is not already larger. The image will
        //  be layed across one or more columns.
        //  Passing EXPAND_COLUMN will result in the width of the column being
        //  increased to accommodate the image if it is not already larger. The image
        //  will be layed across one or many rows.
        //  Passing EXPAND_ROW_AND_COLUMN will result in the height of the row
        //  bing increased along with the width of the column to accomdate the
        //  image if either is not already larger.
        //  Passing OVERLAY_ROW_AND_COLUMN will result in the image being layed
        //  over one or more rows and columns. No row or column will be resized,
        //  instead, code will determine how many rows and columns the image should
        //  overlie.
        public static int EXPAND_ROW = 1;

        public static int EXPAND_COLUMN = 2;

        public static int EXPAND_ROW_AND_COLUMN = 3;

        public static int OVERLAY_ROW_AND_COLUMN = 7;

        //  Modified to support EMU - English Metric Units - used within the OOXML
        //  workbooks, this multoplier is used to convert between measurements in
        //  millimetres and in EMUs
        private static int EMU_PER_MM = 36000;

        public void addImageToSheet(
            (int RowIndex, int ColIndex) cellPosition, 
            ISheet sheet, 
            IDrawing drawing,
            (byte[] ImageBytes, PictureType PictureType) image, 
            double reqImageWidthMM, 
            double reqImageHeightMM, 
            int resizeBehaviour
        )
        {
            //  Convert the string into column and row indices then chain the
            //  call to the overridden addImageToSheet() method.
            CellReference cellRef = new CellReference(cellPosition.RowIndex, cellPosition.ColIndex);
            this.addImageToSheet(cellRef.Col, cellRef.Row, sheet, drawing, image, reqImageWidthMM, reqImageHeightMM, resizeBehaviour);
        }

        public void addImageToSheet(
            int colNumber, 
            int rowNumber, 
            ISheet sheet,
            IDrawing drawing,
            (byte[] ImageBytes, PictureType PictureType) image,
            double reqImageWidthMM, 
            double reqImageHeightMM, 
            int resizeBehaviour
        )
        {
            IClientAnchor anchor;
            ClientAnchorDetail rowClientAnchorDetail;
            ClientAnchorDetail colClientAnchorDetail;
            //PictureType imageType;

            //  Validate the resizeBehaviour parameter.
            if (((resizeBehaviour != AddDimensionedImage.EXPAND_COLUMN)
                        && ((resizeBehaviour != AddDimensionedImage.EXPAND_ROW)
                        && ((resizeBehaviour != AddDimensionedImage.EXPAND_ROW_AND_COLUMN)
                        && (resizeBehaviour != AddDimensionedImage.OVERLAY_ROW_AND_COLUMN)))))
            {
                throw new ArgumentException(("Invalid value passed to the " + "resizeBehaviour parameter of AddDimensionedImage.addImageToSheet()"));
            }

            //  Call methods to calculate how the image and sheet should be
            //  manipulated to accommodate the image; columns and then rows.
            colClientAnchorDetail = this.fitImageToColumns(sheet, colNumber, reqImageWidthMM, resizeBehaviour);
            rowClientAnchorDetail = this.fitImageToRows(sheet, rowNumber, reqImageHeightMM, resizeBehaviour);
            //  Having determined if and how to resize the rows, columns and/or the
            //  image, create the ClientAnchor object to position the image on
            //  the worksheet. Note how the two ClientAnchorDetail records are
            //  interrogated to recover the row/column co-ordinates and any insets.
            //  The first two parameters are not used currently but could be if the
            //  need arose to extend the functionality of this code by adding the
            //  ability to specify that a clear 'border' be placed around the image.
            anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Dx1 = 0;
            anchor.Dy1 = 0;
            if ((colClientAnchorDetail != null))
            {
                anchor.Dx2 = colClientAnchorDetail.getInset();
                anchor.Col1 = colClientAnchorDetail.getFromIndex();
                anchor.Col2 = colClientAnchorDetail.getToIndex();
            }

            if ((rowClientAnchorDetail != null))
            {
                anchor.Dy2 = rowClientAnchorDetail.getInset();
                anchor.Row1 = rowClientAnchorDetail.getFromIndex();
                anchor.Row2 = rowClientAnchorDetail.getToIndex();
            }

            //  For now, set the anchor type to do not move or resize the
            //  image as the size of the row/column is adjusted. This could easily
            //  become another parameter passed to the method. Please read the note
            //  above regarding the behaviour of image resizing.
            anchor.AnchorType = AnchorType.DontMoveAndResize;
            //anchor.AnchorType = AnchorType.MoveAndResize;

            int index = sheet.Workbook.AddPicture(image.ImageBytes, image.PictureType);
            drawing.CreatePicture(anchor, index);
        }

        private ClientAnchorDetail fitImageToColumns(ISheet sheet, int colNumber, double reqImageWidthMM, int resizeBehaviour)
        {
            double colWidthMM;
            double colCoordinatesPerMM;
            int pictureWidthCoordinates;
            ClientAnchorDetail colClientAnchorDetail;

            //  Get the colum's width in millimetres
            colWidthMM = ConvertImageUnits.widthUnits2Millimetres(((short)(sheet.GetColumnWidth(colNumber))));
            //  Check that the column's width will accommodate the image at the
            //  required dimension. If the width of the column is LESS than the
            //  required width of the image, decide how the application should
            //  respond - resize the column or overlay the image across one or more
            //  columns.
            if ((colWidthMM < reqImageWidthMM))
            {
                //  Should the column's width simply be expanded?
                if (((resizeBehaviour == AddDimensionedImage.EXPAND_COLUMN)
                            || (resizeBehaviour == AddDimensionedImage.EXPAND_ROW_AND_COLUMN)))
                {
                    //  Set the width of the column by converting the required image
                    //  width from millimetres into Excel's column width units.
                    sheet.SetColumnWidth(colNumber, ConvertImageUnits.millimetres2WidthUnits(reqImageWidthMM));
                    //  To make the image occupy the full width of the column, convert
                    //  the required width of the image into co-ordinates. This value
                    //  will become the inset for the ClientAnchorDetail class that
                    //  is then instantiated.
                    if (sheet is HSSFSheet)
                    {
                        colWidthMM = reqImageWidthMM;
                        colCoordinatesPerMM = colWidthMM == 0
                            ? 0
                            : ConvertImageUnits.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
                        pictureWidthCoordinates = ((int)((reqImageWidthMM * colCoordinatesPerMM)));
                    }
                    else
                    {
                        pictureWidthCoordinates = (((int)(reqImageWidthMM)) * AddDimensionedImage.EMU_PER_MM);
                    }

                    colClientAnchorDetail = new ClientAnchorDetail(colNumber, colNumber, pictureWidthCoordinates);
                }

                //  If the user has chosen to overlay both rows and columns or just
                //  to expand ONLY the size of the rows, then calculate how to lay
                //  the image out across one or more columns.
                if (((resizeBehaviour == AddDimensionedImage.OVERLAY_ROW_AND_COLUMN)
                            || (resizeBehaviour == AddDimensionedImage.EXPAND_ROW)))
                {
                    colClientAnchorDetail = this.calculateColumnLocation(sheet, colNumber, reqImageWidthMM);
                }

            }

            //  If the column is wider than the image.
            if ((sheet is HSSFSheet))
            {
                //  Mow many co-ordinate positions are there per millimetre?
                colCoordinatesPerMM = (ConvertImageUnits.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM);
                //  Given the width of the image, what should be it's co-ordinate?
                pictureWidthCoordinates = ((int)((reqImageWidthMM * colCoordinatesPerMM)));
            }
            else
            {
                pictureWidthCoordinates = (((int)(reqImageWidthMM)) * AddDimensionedImage.EMU_PER_MM);
            }

            colClientAnchorDetail = new ClientAnchorDetail(colNumber, colNumber, pictureWidthCoordinates);

            return colClientAnchorDetail;
        }

        private ClientAnchorDetail fitImageToRows(ISheet sheet, int rowNumber, double reqImageHeightMM, int resizeBehaviour)
        {
            IRow row;
            double rowHeightMM;
            double rowCoordinatesPerMM;
            int pictureHeightCoordinates;
            ClientAnchorDetail rowClientAnchorDetail;

            //  Get the row and it's height
            row = sheet.GetRow(rowNumber);
            if ((row == null))
            {
                //  Create row if it does not exist.
                row = sheet.CreateRow(rowNumber);
            }

            //  Get the row's height in millimetres
            rowHeightMM = (row.HeightInPoints / ConvertImageUnits.POINTS_PER_MILLIMETRE);
            //  Check that the row's height will accommodate the image at the required
            //  dimensions. If the height of the row is LESS than the required height
            //  of the image, decide how the application should respond - resize the
            //  row or overlay the image across a series of rows.
            if ((rowHeightMM < reqImageHeightMM))
            {
                if (((resizeBehaviour == AddDimensionedImage.EXPAND_ROW)
                            || (resizeBehaviour == AddDimensionedImage.EXPAND_ROW_AND_COLUMN)))
                {
                    row.HeightInPoints = ((float)((reqImageHeightMM * ConvertImageUnits.POINTS_PER_MILLIMETRE)));
                    if ((sheet is HSSFSheet))
                    {
                        rowHeightMM = reqImageHeightMM;
                        rowCoordinatesPerMM = rowHeightMM == 0
                            ? 0
                            : ConvertImageUnits.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;
                        pictureHeightCoordinates = ((int)((reqImageHeightMM * rowCoordinatesPerMM)));
                    }
                    else
                    {
                        pictureHeightCoordinates = ((int)((reqImageHeightMM * AddDimensionedImage.EMU_PER_MM)));
                    }

                    rowClientAnchorDetail = new ClientAnchorDetail(rowNumber, rowNumber, pictureHeightCoordinates);
                }

                //  If the user has chosen to overlay both rows and columns or just
                //  to expand ONLY the size of the columns, then calculate how to lay
                //  the image out ver one or more rows.
                if (((resizeBehaviour == AddDimensionedImage.OVERLAY_ROW_AND_COLUMN)
                            || (resizeBehaviour == AddDimensionedImage.EXPAND_COLUMN)))
                {
                    rowClientAnchorDetail = this.calculateRowLocation(sheet, rowNumber, reqImageHeightMM);
                }

            }

            //  Else, if the image is smaller than the space available
            if ((sheet is HSSFSheet))
            {
                rowCoordinatesPerMM = (ConvertImageUnits.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM);
                pictureHeightCoordinates = ((int)((reqImageHeightMM * rowCoordinatesPerMM)));
            }
            else
            {
                pictureHeightCoordinates = ((int)((reqImageHeightMM * AddDimensionedImage.EMU_PER_MM)));
            }

            rowClientAnchorDetail = new ClientAnchorDetail(rowNumber, rowNumber, pictureHeightCoordinates);

            return rowClientAnchorDetail;
        }

        private ClientAnchorDetail calculateColumnLocation(ISheet sheet, int startingColumn, double reqImageWidthMM)
        {
            ClientAnchorDetail anchorDetail;
            double totalWidthMM = 0;
            double colWidthMM = 0;
            double overlapMM;
            double coordinatePositionsPerMM;
            int toColumn = startingColumn;
            int inset;
            //  Calculate how many columns the image will have to
            //  span in order to be presented at the required size.
            while (totalWidthMM < reqImageWidthMM)
            {
                colWidthMM = ConvertImageUnits.widthUnits2Millimetres(((short)(sheet.GetColumnWidth(toColumn))));
                //  Note use of the cell border width constant. Testing with an image
                //  declared to fit exactly into one column demonstrated that it's
                //  width was greater than the width of the column the POI returned.
                //  Further, this difference was a constant value that I am assuming
                //  related to the cell's borders. Either way, that difference needs
                //  to be allowed for in this calculation.
                totalWidthMM += colWidthMM + ConvertImageUnits.CELL_BORDER_WIDTH_MILLIMETRES;
                toColumn++;
            }

            //  De-crement by one the last column value.
            toColumn--;
            //  Highly unlikely that this will be true but, if the width of a series
            //  of columns is exactly equal to the required width of the image, then
            //  simply build a ClientAnchorDetail object with an inset equal to the
            //  total number of co-ordinate positions available in a column, a
            //  from column co-ordinate (top left hand corner) equal to the value
            //  of the startingColumn parameter and a to column co-ordinate equal
            //  to the toColumn variable.
            // 
            //  Convert both values to ints to perform the test.
            if ((((int)(totalWidthMM)) == ((int)(reqImageWidthMM))))
            {
                //  A problem could occur if the image is sized to fit into one or
                //  more columns. If that occurs, the value in the toColumn variable
                //  will be in error. To overcome this, there are two options, to
                //  ibcrement the toColumn variable's value by one or to pass the
                //  total number of co-ordinate positions to the third paramater
                //  of the ClientAnchorDetail constructor. For no sepcific reason,
                //  the latter option is used below.
                if ((sheet is HSSFSheet))
                {
                    anchorDetail = new ClientAnchorDetail(startingColumn, toColumn, ConvertImageUnits.TOTAL_COLUMN_COORDINATE_POSITIONS);
                }
                else
                {
                    anchorDetail = new ClientAnchorDetail(startingColumn, toColumn, (((int)(reqImageWidthMM)) * AddDimensionedImage.EMU_PER_MM));
                }

            }

            //  In this case, the image will overlap part of another column and it is
            //  necessary to calculate just how much - this will become the inset
            //  for the ClientAnchorDetail object.
            //  Firstly, claculate how much of the image should overlap into
            //  the next column.
            overlapMM = (reqImageWidthMM
                        - (totalWidthMM - colWidthMM));
            //  When the required size is very close indded to the column size,
            //  the calcaulation above can produce a negative value. To prevent
            //  problems occuring in later caculations, this is simply removed
            //  be setting the overlapMM value to zero.
            if ((overlapMM < 0))
            {
                overlapMM = 0;
            }

            if ((sheet is HSSFSheet))
            {
                //  Next, from the columns width, calculate how many co-ordinate
                //  positons there are per millimetre
                coordinatePositionsPerMM = (colWidthMM == 0)
                    ? 0
                    : ConvertImageUnits.TOTAL_COLUMN_COORDINATE_POSITIONS / colWidthMM;
                //  From this figure, determine how many co-ordinat positions to
                //  inset the left hand or bottom edge of the image.
                inset = ((int)((coordinatePositionsPerMM * overlapMM)));
            }
            else
            {
                inset = (((int)(overlapMM)) * AddDimensionedImage.EMU_PER_MM);
            }

            //  Now create the ClientAnchorDetail object, setting the from and to
            //  columns and the inset.
            anchorDetail = new ClientAnchorDetail(startingColumn, toColumn, inset);

            return anchorDetail;
        }

        private ClientAnchorDetail calculateRowLocation(ISheet sheet, int startingRow, double reqImageHeightMM)
        {
            ClientAnchorDetail clientAnchorDetail;
            IRow row;
            double rowHeightMM = 0;
            double totalRowHeightMM = 0;
            double overlapMM;
            double rowCoordinatesPerMM;
            int toRow = startingRow;
            int inset;
            //  Step through the rows in the sheet and accumulate a total of their
            //  heights.
            while (totalRowHeightMM < reqImageHeightMM)
            {
                row = sheet.GetRow(toRow);
                //  Note, if the row does not already exist on the sheet then create
                //  it here.
                if ((row == null))
                {
                    row = sheet.CreateRow(toRow);
                }

                //  Get the row's height in millimetres and add to the running total.
                rowHeightMM = (row.HeightInPoints / ConvertImageUnits.POINTS_PER_MILLIMETRE);
                totalRowHeightMM += rowHeightMM;
                toRow++;
            }

            //  Owing to the way the loop above works, the rowNumber will have been
            //  incremented one row too far. Undo that here.
            toRow--;
            //  Check to see whether the image should occupy an exact number of
            //  rows. If so, build the ClientAnchorDetail record to point
            //  to those rows and with an inset of the total number of co-ordinate
            //  position in the row.
            // 
            //  To overcome problems that can occur with comparing double values for
            //  equality, cast both to int(s) to truncate the value; VERY crude and
            //  I do not really like it!!
            if ((((int)(totalRowHeightMM)) == ((int)(reqImageHeightMM))))
            {
                if ((sheet is HSSFSheet))
                {
                    clientAnchorDetail = new ClientAnchorDetail(startingRow, toRow, ConvertImageUnits.TOTAL_ROW_COORDINATE_POSITIONS);
                }
                else
                {
                    clientAnchorDetail = new ClientAnchorDetail(startingRow, toRow, (((int)(reqImageHeightMM)) * AddDimensionedImage.EMU_PER_MM));
                }

            }
            else
            {
                //  Calculate how far the image will project into the next row. Note
                //  that the height of the last row assessed is subtracted from the
                //  total height of all rows assessed so far.
                overlapMM = (reqImageHeightMM - (totalRowHeightMM - rowHeightMM));

                //  To prevent an exception being thrown when the required width of
                //  the image is very close indeed to the column size.
                if ((overlapMM < 0))
                {
                    overlapMM = 0;
                }

                if ((sheet is HSSFSheet))
                {
                    rowCoordinatesPerMM = rowHeightMM == 0 
                        ? 0
                        : ConvertImageUnits.TOTAL_ROW_COORDINATE_POSITIONS / rowHeightMM;
                    inset = ((int)((overlapMM * rowCoordinatesPerMM)));
                }
                else
                {
                    inset = (((int)(overlapMM)) * AddDimensionedImage.EMU_PER_MM);
                }

                clientAnchorDetail = new ClientAnchorDetail(startingRow, toRow, inset);
            }

            return clientAnchorDetail;
        }

        //public static void main(String[] args)
        //{
        //    if ((args.length < 2))
        //    {
        //        System.err.println("Usage: AddDimensionedImage imageFile outputFile");
        //        return;
        //    }

        //    string imageFile = args[0];
        //    string outputFile = args[1];
        //    try
        //    {
        //    }


        //Workbook workbook = new HSSFWorkbook();
        //    FileOutputStream fos = new FileOutputStream(outputFile);
        //    //  OR XSSFWorkbook
        //    ISheet sheet = workbook.createSheet("Picture Test");
        //    (new AddDimensionedImage() + this.addImageToSheet("B5", sheet, sheet.createDrawingPatriarch(), (new File(imageFile) + toURI().toURL()), 100, 40, AddDimensionedImage.EXPAND_ROW_AND_COLUMN));
        //    workbook.write(fos);
        //}

        public class ClientAnchorDetail
        {

            private int fromIndex;

            private int toIndex;

            private int inset;

            public ClientAnchorDetail(int fromIndex, int toIndex, int inset)
            {
                this.fromIndex = fromIndex;
                this.toIndex = toIndex;
                this.inset = inset;
            }

            public int getFromIndex()
            {
                return this.fromIndex;
            }

            public int getToIndex()
            {
                return this.toIndex;
            }

            public int getInset()
            {
                return this.inset;
            }
        }

        public class ConvertImageUnits
        {

            //  Each cell conatins a fixed number of co-ordinate points; this number
            //  does not vary with row height or column width or with font. These two
            //  constants are defined below.
            public static int TOTAL_COLUMN_COORDINATE_POSITIONS = 1023;
            public static int TOTAL_ROW_COORDINATE_POSITIONS = 255;

            //  The resoultion of an image can be expressed as a specific number
            //  of pixels per inch. Displays and printers differ but 96 pixels per
            //  inch is an acceptable standard to beging with.
            public static int PIXELS_PER_INCH = 96;

            //  Cnstants that defines how many pixels and points there are in a
            //  millimetre. These values are required for the conversion algorithm.
            public static double PIXELS_PER_MILLIMETRES = 3.78;
            public static double POINTS_PER_MILLIMETRE = 2.83;

            //  The column width returned by HSSF and the width of a picture when
            //  positioned to exactly cover one cell are different by almost exactly
            //  2mm - give or take rounding errors. This constant allows that
            //  additional amount to be accounted for when calculating how many
            //  celles the image ought to overlie.
            public static double CELL_BORDER_WIDTH_MILLIMETRES = 2;
            public static short EXCEL_COLUMN_WIDTH_FACTOR = 256;
            public static int UNIT_OFFSET_LENGTH = 7;
            private static short[] UNIT_OFFSET_MAP = { 0, 36, 73, 109, 146, 182, 219 };

        
        public static short pixel2WidthUnits(int pxs)
            {
                short widthUnits = ((short)((EXCEL_COLUMN_WIDTH_FACTOR
                            * (pxs / UNIT_OFFSET_LENGTH))));
                widthUnits += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)];
                return widthUnits;
            }

            public static int widthUnits2Pixel(short widthUnits)
            {
                int pixels = ((widthUnits / EXCEL_COLUMN_WIDTH_FACTOR)
                            * UNIT_OFFSET_LENGTH);
                int offsetWidthUnits = (widthUnits % EXCEL_COLUMN_WIDTH_FACTOR);
                pixels += (int) Math.Round((offsetWidthUnits
                                / (((float)(EXCEL_COLUMN_WIDTH_FACTOR)) / UNIT_OFFSET_LENGTH)));
                return pixels;
            }

            public static double widthUnits2Millimetres(short widthUnits)
            {
                return (ConvertImageUnits.widthUnits2Pixel(widthUnits) / ConvertImageUnits.PIXELS_PER_MILLIMETRES);
            }

            public static int millimetres2WidthUnits(double millimetres)
            {
                return ConvertImageUnits.pixel2WidthUnits(((int)((millimetres * ConvertImageUnits.PIXELS_PER_MILLIMETRES))));
            }
        }
    }
}
