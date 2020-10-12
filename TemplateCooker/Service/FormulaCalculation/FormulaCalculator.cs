using ClosedXML.Excel;
using System;
using System.Linq;

namespace TemplateCooker.Service.FormulaCalculation
{
    public class FormulaCalculator
    {
        private readonly FormulaCalculationOptions _options;

        public FormulaCalculator(FormulaCalculationOptions options)
        {
            _options = options;
        }

        public void Recalculate(IXLWorkbook workbook)
        {
            workbook.RecalculateAllFormulas();
        }
    }
}