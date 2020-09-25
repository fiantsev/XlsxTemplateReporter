using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Test_NpoiDotnet.Operators
{
    public class RangeCopier
    {
        public void CopyRange(XSSFSheet sheetFrom, XSSFSheet sheetTo, CellRangeAddress from, CellRangeAddress to)
        {
            InnerCopyRange(sheetFrom, sheetTo, from, to);
        }

        private void InnerCopyRange(ISheet sourceSheet, ISheet destinationSheet, CellRangeAddress rangeFrom, CellRangeAddress rangeTo)
        {
            var rangeToFirstRow = rangeTo.FirstRow;
            var rangeToFirstColumn = rangeTo.FirstColumn;

            for (var rowNum = rangeFrom.FirstRow; rowNum <= rangeFrom.LastRow; rowNum++)
            {
                IRow sourceRow = sourceSheet.GetRow(rowNum);

                if (destinationSheet.GetRow(rowNum) == null)
                    destinationSheet.CreateRow(rowNum);

                if (sourceRow != null)
                {
                    var destinationRowIndex = rangeToFirstRow + (rowNum - rangeFrom.FirstRow);
                    IRow destinationRow = destinationSheet.GetRow(destinationRowIndex)
                        ?? destinationSheet.CreateRow(destinationRowIndex);

                    for (var col = rangeFrom.FirstColumn; col < sourceRow.LastCellNum && col <= rangeFrom.LastColumn; col++)
                    {
                        var destinationCellIndex = rangeToFirstColumn + (col - rangeFrom.FirstColumn);
                        var destinationCell = destinationRow.GetCell(destinationCellIndex)
                            ?? destinationRow.CreateCell(destinationCellIndex);

                        CopyCell(sourceRow.GetCell(col), destinationCell);
                    }
                }
            }
        }

        private void CopyCell(ICell source, ICell destination)
        {
            if (destination != null && source != null)
            {
                //you can comment these out if you don't want to copy the style ...
                destination.CellComment = source.CellComment;
                destination.CellStyle = source.CellStyle;
                destination.Hyperlink = source.Hyperlink;

                switch (source.CellType)
                {
                    case CellType.Formula:
                        destination.CellFormula = source.CellFormula; break;
                    case CellType.Numeric:
                        destination.SetCellValue(source.NumericCellValue); break;
                    case CellType.String:
                        destination.SetCellValue(source.StringCellValue); break;
                }
            }
        }
    }
}
