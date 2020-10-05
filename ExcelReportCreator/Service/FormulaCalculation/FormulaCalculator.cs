using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelReportCreatorProject.Service.FormulaCalculation
{
    public class FormulaCalculator
    {
        private readonly FormulaCalculatorOptions _options;

        public FormulaCalculator(FormulaCalculatorOptions options)
        {
            _options = options;
        }

        public void Recalculate(Stream workbookStream)
        {
            IWorkbook workbook = new XSSFWorkbook(workbookStream);

            var formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            foreach (var sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
            {
                var sheet = workbook.GetSheetAt(sheetIndex);

                for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; ++rowIndex)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    for (var cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; ++cellIndex)
                    {
                        var cell = row.GetCell(cellIndex);
                        if (cell == null) continue;

                        formulaEvaluator.EvaluateFormulaCell(cell);
                    }
                }
            }

            workbook.Write(workbookStream);
        }
    }
}