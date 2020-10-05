using System.IO;
using ExcelReportCreatorProject.Service.Creation;
using ExcelReportCreatorProject.Service.FormulaCalculation;

namespace ExcelReportCreatorProject
{
    public class TemplateBuilder
    {
        private readonly Stream _template;

        public TemplateBuilder(Stream template)
        {
            _template = template;
            //template.CopyTo(_template);
        }

        public TemplateBuilder InjectData(DocumentInjectorOptions options) {
            var documentInjector = new DocumentInjector(options);
            documentInjector.Inject(_template);
            return this;
        }

        public TemplateBuilder RecalculateFormulas(FormulaCalculatorOptions options)
        {
            var formulaEvaluator = new FormulaCalculator(options);
            formulaEvaluator.Recalculate(_template);
            return this;
        }

        //public Stream Build()
        //{
        //    return _template;
        //}
    }
}