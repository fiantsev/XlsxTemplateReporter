using System.Collections;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelReportCreatorProject.Domain.Npoi
{
    public class SheetCollection : IEnumerable<ISheet>
    {
        private readonly IWorkbook _workbook;

        public SheetCollection(IWorkbook workbook) {
            _workbook = workbook;
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            foreach (var sheetIndex in Enumerable.Range(0, _workbook.NumberOfSheets))
                yield return _workbook.GetSheetAt(sheetIndex);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
