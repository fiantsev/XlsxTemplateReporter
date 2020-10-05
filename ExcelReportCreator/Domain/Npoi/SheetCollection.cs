using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelReportCreatorProject.Domain.Npoi
{
    public class SheetCollection : IEnumerable<IXLWorksheet>
    {
        private readonly XLWorkbook _workbook;

        public SheetCollection(XLWorkbook workbook) {
            _workbook = workbook;
        }

        public IEnumerator<IXLWorksheet> GetEnumerator()
        {
            return _workbook.Worksheets.GetEnumerator();
            //foreach (var sheetIndex in Enumerable.Range(0, _workbook.NumberOfSheets))
            //    yield return _workbook.GetSheetAt(sheetIndex);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
