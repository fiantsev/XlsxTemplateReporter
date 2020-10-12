using ClosedXML.Excel;
using System;
using System.Linq;

namespace TemplateCooker.Service.FormulaCalculation
{
    public class FormulaCalculator
    {
        private readonly FormulaCalculatorOptions _options;

        public FormulaCalculator(FormulaCalculatorOptions options)
        {
            _options = options;
        }

        public void Recalculate(IXLWorkbook workbook)
        {
            workbook.RecalculateAllFormulas();
        }
    }
}