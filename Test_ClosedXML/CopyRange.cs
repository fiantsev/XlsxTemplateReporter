using ClosedXML.Excel;

namespace Test_ClosedXML
{
    public class CopyRange
    {
        public void Copy(IXLRange from, IXLRange to)
        {
            InnerCopyRange(from, to);
        }

        private void InnerCopyRange(IXLRange from, IXLRange to)
        {
            //var workbook = new XLWorkbook("BasicTable.xlsx");
            //var ws = workbook.Worksheet(1);

            // Define a range with the data
            //var firstTableCell = ws.FirstCellUsed();
            //var lastTableCell = ws.LastCellUsed();
            //var rngData = sheetFrom.Range(from.RangeAddress, to);

            // Copy the table to another worksheet
            //var wsCopy = workbook.Worksheets.Add("Contacts Copy");
            //sheetTo.Cell(to.FirstCellUsed().Address.RowNumber, to.FirstCellUsed().Address.ColumnNumber).Value = from;
            to.FirstCellUsed().Value = from;

            //workbook.SaveAs("CopyingRanges.xlsx");
        }
    }
}
